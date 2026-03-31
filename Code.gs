/**
 * Grafik Foto Schedule Automation
 *
 * Быстрая навигация по проекту:
 * - onOpen: добавляет меню в Google Sheets
 * - showCreateMonthDialog: открывает окно генерации
 * - getDialogData: подготавливает данные для диалога
 * - createMonthSchedule: главный сценарий создания нового месяца
 * - solveSchedule_ / solveWithSelectedPairs_: подбирают допустимый график
 * - validateSolution_: формирует проверки для отчёта после генерации
 *
 * Главная идея:
 * 1. Берём предыдущий месячный лист как шаблон
 * 2. Читаем постоянные правила сотрудников с листа "Настройки"
 * 3. Получаем от пользователя месяц, нормы часов и отпуск
 * 4. Собираем допустимый график по ограничениям
 * 5. Копируем шаблон, заполняем новый лист и показываем отчёт
 */

const APP_CONFIG = {
  settingsSheetName: 'Настройки',
  settingsHeaders: [
    'ФИО',
    'Активен',
    'Часы смены',
    'Фиксированные выходные',
    'Требуется 1 пара сб+вс',
    'Все четверги рабочие',
    'Участвует в дежурствах',
  ],
  sampleLegendStartColumn: 10,
  sampleLegendHeaders: ['Обычная смена', 'Выходной', 'Отпуск', 'Дежурство'],
  dataStartRow: 3,
  dayStartColumn: 5,
  maxDayColumns: 31,
  monthNames: ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'],
  weekdayNames: ['вс', 'пн', 'вт', 'ср', 'чт', 'пт', 'сб'],
  defaultShiftHours: 8,
  offCodeFallback: 'в',
  vacationCodeFallback: 'о',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('График фотографов')
    .addItem('Создать месяц', 'showCreateMonthDialog')
    .addToUi();
}

function showCreateMonthDialog() {
  const html = HtmlService.createHtmlOutputFromFile('GenerateMonthDialog')
    .setWidth(960)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Создать график нового месяца');
}

function getDialogData() {
  const spreadsheet = SpreadsheetApp.getActive();
  const templateSheet = findBestTemplateSheet_(spreadsheet);
  if (!templateSheet) {
    throw new Error('Не найден ни один месячный лист. Добавьте лист месяца вроде "мар 2026" и повторите.');
  }

  const layout = getMonthSheetLayout_(templateSheet);
  const settingsSheet = ensureSettingsSheet_(spreadsheet, layout);
  const settingsMap = getSettingsMap_(settingsSheet);
  const suggested = getSuggestedTargetMonth_(templateSheet.getName());
  const norms = templateSheet
    .getRange(APP_CONFIG.dataStartRow, 3, Math.max(layout.employeeRows.length, 1), 1)
    .getDisplayValues()
    .map((row) => row[0]);

  const employees = layout.employeeRows
    .map((entry, index) => {
      const settings = settingsMap[entry.name] || createDefaultSettingsRow_(entry.name);
      if (!settings.active) {
        return null;
      }

      return {
        name: entry.name,
        defaultNormHours: normalizeNormForDialog_(norms[index]),
        shiftHours: settings.shiftHours,
        fixedOffWeekdays: settings.fixedOffWeekdaysRaw,
        requiresWeekendPair: settings.requiresWeekendPair,
        allThursdaysWork: settings.allThursdaysWork,
        dutyParticipation: settings.participatesInDuty,
      };
    })
    .filter(Boolean);

  return {
    monthNames: APP_CONFIG.monthNames,
    suggestedMonth: suggested.month + 1,
    suggestedYear: suggested.year,
    employees,
    settingsSheetName: APP_CONFIG.settingsSheetName,
  };
}

function createMonthSchedule(payload) {
  const spreadsheet = SpreadsheetApp.getActive();
  const input = normalizeDialogPayload_(payload);
  const templateSheet = findTemplateSheetForTarget_(spreadsheet, input.year, input.monthIndex);
  if (!templateSheet) {
    throw new Error('Не удалось подобрать шаблонный месячный лист.');
  }

  const layout = getMonthSheetLayout_(templateSheet);
  const settingsSheet = ensureSettingsSheet_(spreadsheet, layout);
  const settingsMap = getSettingsMap_(settingsSheet);
  const employeeRowsByName = {};
  layout.employeeRows.forEach((entry) => {
    employeeRowsByName[entry.name] = entry.row;
  });

  const activeEmployees = buildEmployeeModels_(layout, employeeRowsByName, settingsMap, input);
  if (!activeEmployees.length) {
    throw new Error('Не найдено ни одного активного сотрудника для генерации графика.');
  }

  const templateFormats = detectFormatSamplesFromSheet_(templateSheet, layout);
  refreshFormatSamples_(settingsSheet, templateFormats);
  const formatSources = getFormatSampleRanges_(settingsSheet, templateFormats);

  const days = buildMonthDays_(input.year, input.monthIndex);
  const solution = solveSchedule_(activeEmployees, days);
  const targetName = buildMonthSheetName_(input.year, input.monthIndex);
  const outputSheet = copySheetWithUniqueName_(spreadsheet, templateSheet, targetName);

  applyScheduleToSheet_(outputSheet, layout, days, activeEmployees, solution, formatSources);
  const report = buildGenerationReport_(activeEmployees, days, solution, outputSheet.getName());

  return {
    sheetName: outputSheet.getName(),
    employeeCount: activeEmployees.length,
    daysInMonth: days.length,
    reportTitle: report.title,
    reportText: report.text,
  };
}

function buildGenerationReport_(employees, days, solution, sheetName) {
  const lines = [];
  const checks = validateSolution_(employees, days, solution);
  const dutyStats = summarizeDutyStats_(employees, days, solution);
  const vacationSummary = summarizeVacations_(employees);

  lines.push('Лист: ' + sheetName);
  lines.push('');
  lines.push('Дежурства:');
  employees.forEach((employee) => {
    const stats = dutyStats[employee.id];
    lines.push(
      '- ' +
        employee.name +
        ': будни ' +
        stats.weekday +
        ', выходные ' +
        stats.weekend +
        ', всего ' +
        (stats.weekday + stats.weekend)
    );
  });

  lines.push('');
  lines.push('Отпуска:');
  if (vacationSummary.length) {
    vacationSummary.forEach((line) => lines.push('- ' + line));
  } else {
    lines.push('- отпусков в выбранном месяце нет');
  }

  lines.push('');
  lines.push('Проверка правил:');
  checks.forEach((check) => {
    lines.push('- ' + (check.ok ? 'OK' : 'Ошибка') + ': ' + check.label);
  });

  return {
    title: 'Проверка после генерации',
    text: lines.join('\n'),
  };
}

function summarizeDutyStats_(employees, days, solution) {
  const stats = {};
  employees.forEach((employee) => {
    stats[employee.id] = { weekday: 0, weekend: 0 };
  });

  Object.keys(solution.duties).forEach((key) => {
    const parts = key.split(':');
    const employeeId = parts[0];
    const dayIndex = Number(parts[1]);
    if (!stats[employeeId]) {
      return;
    }

    if (days[dayIndex].isWeekday) {
      stats[employeeId].weekday += 1;
    } else {
      stats[employeeId].weekend += 1;
    }
  });

  return stats;
}

function summarizeVacations_(employees) {
  return employees
    .map((employee) => {
      const daysList = Array.from(employee.vacationDays).sort((left, right) => left - right);
      if (!daysList.length) {
        return '';
      }

      return employee.name + ': ' + formatDayRanges_(daysList);
    })
    .filter(Boolean);
}

function validateSolution_(employees, days, solution) {
  const checks = [];
  const dutyStats = summarizeDutyStats_(employees, days, solution);

  const emptyDays = [];
  for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
    const workers = employees.filter((employee) => solution.statuses[employee.id][dayIndex] === 'WORK').length;
    if (workers < 1) {
      emptyDays.push(days[dayIndex].day);
    }
  }
  checks.push({
    ok: emptyDays.length === 0,
    label: emptyDays.length === 0 ? 'каждый день закрыт хотя бы одним сотрудником' : 'пустые дни: ' + emptyDays.join(', '),
  });

  const badDutyDays = [];
  for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
    const count = employees.filter((employee) => solution.duties[employee.id + ':' + dayIndex]).length;
    if (count !== 1) {
      badDutyDays.push(days[dayIndex].day);
    }
  }
  checks.push({
    ok: badDutyDays.length === 0,
    label: badDutyDays.length === 0 ? 'ровно один дежурный на каждый день' : 'дни с неверным числом дежурных: ' + badDutyDays.join(', '),
  });

  employees.forEach((employee) => {
    const workDays = solution.statuses[employee.id].filter((status) => status === 'WORK').length;
    checks.push({
      ok: workDays === employee.targetWorkDays,
      label:
        employee.name +
        ': норма ' +
        employee.targetWorkDays * employee.shiftHours +
        ' ч, факт ' +
        workDays * employee.shiftHours +
        ' ч',
    });

    const offRunViolation = getOffRunViolation_(solution.statuses[employee.id], days, 3);
    checks.push({
      ok: !offRunViolation,
      label: !offRunViolation
        ? employee.name + ': не более 3 выходных подряд'
        : employee.name +
          ': серия выходных ' +
          offRunViolation.length +
          ' дн. подряд (' +
          formatDayRange_(offRunViolation.startDay, offRunViolation.endDay) +
          ')',
    });

    if (employee.requiresWeekendPair) {
      let hasPair = false;
      for (let dayIndex = 0; dayIndex < days.length - 1; dayIndex += 1) {
        if (
          days[dayIndex].weekday === 6 &&
          days[dayIndex + 1].weekday === 0 &&
          solution.statuses[employee.id][dayIndex] === 'OFF' &&
          solution.statuses[employee.id][dayIndex + 1] === 'OFF'
        ) {
          hasPair = true;
          break;
        }
      }

      checks.push({
        ok: hasPair,
        label: employee.name + ': есть хотя бы одна пара сб+вс',
      });
    }

    if (employee.fixedOffWeekdays.size) {
      const violations = [];
      for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
        if (
          employee.fixedOffWeekdays.has(days[dayIndex].weekday) &&
          solution.statuses[employee.id][dayIndex] === 'WORK'
        ) {
          violations.push(days[dayIndex].day);
        }
      }

      checks.push({
        ok: violations.length === 0,
        label:
          violations.length === 0
            ? employee.name + ': фиксированные выходные соблюдены'
            : employee.name + ': нарушены фиксированные выходные в днях ' + violations.join(', '),
      });
    }

    if (employee.allThursdaysWork) {
      const violations = [];
      for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
        if (
          days[dayIndex].isThursday &&
          !employee.vacationDays.has(days[dayIndex].day) &&
          solution.statuses[employee.id][dayIndex] !== 'WORK'
        ) {
          violations.push(days[dayIndex].day);
        }
      }

      checks.push({
        ok: violations.length === 0,
        label:
          violations.length === 0
            ? employee.name + ': все четверги рабочие'
            : employee.name + ': нарушены четверги в днях ' + violations.join(', '),
      });
    }
  });

  const weekdayCounts = employees.map((employee) => dutyStats[employee.id].weekday);
  const weekdaySpread = Math.max.apply(null, weekdayCounts) - Math.min.apply(null, weekdayCounts);
  checks.push({
    ok: weekdaySpread <= 1,
    label: 'разница по будничным дежурствам: ' + weekdaySpread,
  });

  return checks;
}

function formatDayRanges_(daysList) {
  const parts = [];
  let start = daysList[0];
  let previous = daysList[0];

  for (let index = 1; index <= daysList.length; index += 1) {
    const current = daysList[index];
    if (current === previous + 1) {
      previous = current;
      continue;
    }

    parts.push(start === previous ? String(start) : start + '-' + previous);
    start = current;
    previous = current;
  }

  return parts.join(', ');
}

function formatDayRange_(startDay, endDay) {
  return startDay === endDay ? String(startDay) : startDay + '-' + endDay;
}

function getOffRunViolation_(statuses, days, maxLength) {
  let runStartIndex = -1;
  let runLength = 0;

  for (let index = 0; index <= statuses.length; index += 1) {
    const status = index < statuses.length ? statuses[index] : null;
    if (status === 'OFF') {
      if (runStartIndex === -1) {
        runStartIndex = index;
      }
      runLength += 1;
      continue;
    }

    if (runLength > maxLength) {
      return {
        startDay: days[runStartIndex].day,
        endDay: days[index - 1].day,
        length: runLength,
      };
    }

    runStartIndex = -1;
    runLength = 0;
  }

  return null;
}

function ensureSettingsSheet_(spreadsheet, layout) {
  let sheet = spreadsheet.getSheetByName(APP_CONFIG.settingsSheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(APP_CONFIG.settingsSheetName);
  }

  const headerRange = sheet.getRange(1, 1, 1, APP_CONFIG.settingsHeaders.length);
  headerRange.setValues([APP_CONFIG.settingsHeaders]);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);

  const legendHeaderRange = sheet.getRange(1, APP_CONFIG.sampleLegendStartColumn, 1, APP_CONFIG.sampleLegendHeaders.length);
  legendHeaderRange.setValues([APP_CONFIG.sampleLegendHeaders]);
  legendHeaderRange.setFontWeight('bold');

  const noteRange = sheet.getRange(3, APP_CONFIG.sampleLegendStartColumn, 1, 4);
  noteRange.breakApart();
  const noteCell = noteRange;
  noteCell.merge();
  noteCell.setValue('Образцы форматов используются генератором. Их можно обновлять из текущего шаблонного месяца.');
  noteCell.setWrap(true);
  noteCell.setVerticalAlignment('top');

  syncSettingsEmployees_(sheet, layout.employeeRows.map((entry) => entry.name));
  return sheet;
}

function syncSettingsEmployees_(settingsSheet, employeeNames) {
  const settingsMap = getSettingsMap_(settingsSheet);
  const missing = employeeNames.filter((name) => !settingsMap[name]);
  if (!missing.length) {
    return;
  }

  const nameColumnHeight = Math.max(settingsSheet.getLastRow() - 1, 1);
  const nameValues = settingsSheet.getRange(2, 1, nameColumnHeight, 1).getDisplayValues();
  let lastNameRow = 1;
  nameValues.forEach((row, index) => {
    if (String(row[0] || '').trim()) {
      lastNameRow = index + 2;
    }
  });

  const startRow = Math.max(lastNameRow + 1, 2);
  const values = missing.map((name) => {
    const row = createDefaultSettingsRow_(name);
    return [
      row.name,
      row.active,
      row.shiftHours,
      row.fixedOffWeekdaysRaw,
      row.requiresWeekendPair,
      row.allThursdaysWork,
      row.participatesInDuty,
    ];
  });

  settingsSheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
}

function createDefaultSettingsRow_(name) {
  return {
    name,
    active: true,
    shiftHours: APP_CONFIG.defaultShiftHours,
    fixedOffWeekdays: new Set(),
    fixedOffWeekdaysRaw: '',
    requiresWeekendPair: false,
    allThursdaysWork: true,
    participatesInDuty: true,
  };
}

function getSettingsMap_(settingsSheet) {
  const lastRow = settingsSheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }

  const values = settingsSheet.getRange(2, 1, lastRow - 1, APP_CONFIG.settingsHeaders.length).getValues();
  const map = {};

  values.forEach((row) => {
    const name = String(row[0] || '').trim();
    if (!name) {
      return;
    }

    map[name] = {
      name,
      active: asBoolean_(row[1], true),
      shiftHours: normalizeShiftHours_(row[2]),
      fixedOffWeekdays: parseFixedOffWeekdays_(row[3]),
      fixedOffWeekdaysRaw: String(row[3] || '').trim(),
      requiresWeekendPair: asBoolean_(row[4], false),
      allThursdaysWork: asBoolean_(row[5], true),
      participatesInDuty: asBoolean_(row[6], true),
    };
  });

  return map;
}

function buildEmployeeModels_(layout, employeeRowsByName, settingsMap, input) {
  const inputByName = {};
  input.employees.forEach((entry) => {
    inputByName[entry.name] = entry;
  });

  return layout.employeeRows
    .map((entry) => {
      const settings = settingsMap[entry.name] || createDefaultSettingsRow_(entry.name);
      if (!settings.active) {
        return null;
      }

      const employeeInput = inputByName[entry.name];
      if (!employeeInput) {
        throw new Error('Для сотрудника "' + entry.name + '" не передана норма часов.');
      }

      const shiftHours = settings.shiftHours;
      if (shiftHours !== APP_CONFIG.defaultShiftHours) {
        throw new Error('Пока поддерживаются только смены по 8 часов. Проверьте лист "Настройки" для "' + entry.name + '".');
      }

      if (employeeInput.normHours % shiftHours !== 0) {
        throw new Error('Норма часов для "' + entry.name + '" должна быть кратна ' + shiftHours + '.');
      }

      const vacationDays = buildVacationDaySet_(
        employeeInput.vacationFrom,
        employeeInput.vacationTo,
        input.year,
        input.monthIndex
      );

      return {
        id: entry.name,
        name: entry.name,
        row: employeeRowsByName[entry.name],
        shiftHours,
        targetWorkDays: employeeInput.normHours / shiftHours,
        fixedOffWeekdays: settings.fixedOffWeekdays,
        requiresWeekendPair: settings.requiresWeekendPair,
        allThursdaysWork: settings.allThursdaysWork,
        participatesInDuty: settings.participatesInDuty,
        vacationDays,
      };
    })
    .filter(Boolean);
}

function solveSchedule_(employees, days) {
  const employeesNeedingPair = employees.filter((employee) => employee.requiresWeekendPair);
  const pairOptions = {};

  employeesNeedingPair.forEach((employee) => {
    const options = getWeekendPairOptions_(employee, days);
    if (!options.length) {
      throw new Error('Для сотрудника "' + employee.name + '" не найдено ни одной доступной пары сб+вс без отпуска.');
    }

    pairOptions[employee.id] = options;
  });

  let foundSolution = null;

  function searchPairs(index, selectedPairs) {
    if (foundSolution) {
      return;
    }

    if (index >= employeesNeedingPair.length) {
      const candidate = solveWithSelectedPairs_(employees, days, selectedPairs);
      if (candidate) {
        foundSolution = candidate;
      }
      return;
    }

    const employee = employeesNeedingPair[index];
    pairOptions[employee.id].forEach((pair) => {
      const nextSelection = Object.assign({}, selectedPairs);
      nextSelection[employee.id] = pair;
      searchPairs(index + 1, nextSelection);
    });
  }

  searchPairs(0, {});

  if (!foundSolution) {
    throw new Error('Не удалось построить график с текущими ограничениями. Проверьте нормы часов, отпуска, настройки сотрудников и правило "не больше 3 выходных подряд".');
  }

  return foundSolution;
}

function solveWithSelectedPairs_(employees, days, selectedPairs) {
  const statuses = {};
  const protectedWork = {};
  const coverage = new Array(days.length).fill(0);

  employees.forEach((employee) => {
    statuses[employee.id] = new Array(days.length).fill(null);
    protectedWork[employee.id] = new Array(days.length).fill(false);
  });

  for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
    const day = days[dayIndex];

    for (let employeeIndex = 0; employeeIndex < employees.length; employeeIndex += 1) {
      const employee = employees[employeeIndex];
      if (employee.vacationDays.has(day.day)) {
        if (!setStatus_(statuses[employee.id], dayIndex, 'VACATION')) {
          return null;
        }
        continue;
      }

      if (employee.fixedOffWeekdays.has(day.weekday)) {
        if (!setStatus_(statuses[employee.id], dayIndex, 'OFF')) {
          return null;
        }
      }

      if (employee.allThursdaysWork && day.isThursday) {
        if (!setStatus_(statuses[employee.id], dayIndex, 'WORK')) {
          return null;
        }
        protectedWork[employee.id][dayIndex] = true;
      }
    }
  }

  for (let employeeIndex = 0; employeeIndex < employees.length; employeeIndex += 1) {
    const employee = employees[employeeIndex];
    const pair = selectedPairs[employee.id];
    if (!pair) {
      continue;
    }

    if (!setStatus_(statuses[employee.id], pair.saturdayIndex, 'OFF')) {
      return null;
    }

    if (!setStatus_(statuses[employee.id], pair.sundayIndex, 'OFF')) {
      return null;
    }
  }

  employees.forEach((employee) => {
    for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
      if (!statuses[employee.id][dayIndex]) {
        statuses[employee.id][dayIndex] = 'WORK';
      }
      if (statuses[employee.id][dayIndex] === 'WORK') {
        coverage[dayIndex] += 1;
      }
    }
  });

  const excessByEmployee = {};
  for (let employeeIndex = 0; employeeIndex < employees.length; employeeIndex += 1) {
    const employee = employees[employeeIndex];
    const baselineWorkDays = statuses[employee.id].filter((value) => value === 'WORK').length;
    const excess = baselineWorkDays - employee.targetWorkDays;
    if (excess < 0) {
      return null;
    }
    excessByEmployee[employee.id] = excess;
  }

  const employeesToTrim = employees
    .filter((employee) => excessByEmployee[employee.id] > 0)
    .sort((left, right) => excessByEmployee[right.id] - excessByEmployee[left.id]);

  function trimEmployee(index) {
    if (index >= employeesToTrim.length) {
      return true;
    }

    const employee = employeesToTrim[index];
    const needToRemove = excessByEmployee[employee.id];
    const candidates = getRemovableWorkDays_(employee, statuses[employee.id], protectedWork[employee.id], coverage, days);

    if (candidates.length < needToRemove) {
      return false;
    }

    function chooseDays(candidateIndex, remaining) {
      if (remaining === 0) {
        return trimEmployee(index + 1);
      }

      if (candidates.length - candidateIndex < remaining) {
        return false;
      }

      for (let cursor = candidateIndex; cursor < candidates.length; cursor += 1) {
        const dayIndex = candidates[cursor];
        if (coverage[dayIndex] <= 1) {
          continue;
        }

        statuses[employee.id][dayIndex] = 'OFF';
        coverage[dayIndex] -= 1;

        if (getOffRunViolation_(statuses[employee.id], days, 3)) {
          coverage[dayIndex] += 1;
          statuses[employee.id][dayIndex] = 'WORK';
          continue;
        }

        if (chooseDays(cursor + 1, remaining - 1)) {
          return true;
        }

        coverage[dayIndex] += 1;
        statuses[employee.id][dayIndex] = 'WORK';
      }

      return false;
    }

    return chooseDays(0, needToRemove);
  }

  if (!trimEmployee(0)) {
    return null;
  }

  for (let employeeIndex = 0; employeeIndex < employees.length; employeeIndex += 1) {
    const employee = employees[employeeIndex];
    const actualWorkDays = statuses[employee.id].filter((value) => value === 'WORK').length;
    if (actualWorkDays !== employee.targetWorkDays) {
      return null;
    }

    if (getOffRunViolation_(statuses[employee.id], days, 3)) {
      return null;
    }
  }

  for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
    if (coverage[dayIndex] < 1) {
      return null;
    }
  }

  const duties = assignDuties_(employees, days, statuses);
  if (!duties) {
    return null;
  }

  return {
    statuses,
    duties,
    selectedPairs,
  };
}

function assignDuties_(employees, days, statuses) {
  const weekdayDutyCounts = {};
  const weekendDutyCounts = {};
  const duties = {};
  employees.forEach((employee) => {
    weekdayDutyCounts[employee.id] = 0;
    weekendDutyCounts[employee.id] = 0;
  });

  for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
    const day = days[dayIndex];

    const eligible = employees
      .filter((employee) => employee.participatesInDuty && statuses[employee.id][dayIndex] === 'WORK')
      .sort((left, right) => {
        const weekdayDelta = weekdayDutyCounts[left.id] - weekdayDutyCounts[right.id];
        if (weekdayDelta !== 0) {
          return weekdayDelta;
        }

        const weekendDelta = weekendDutyCounts[left.id] - weekendDutyCounts[right.id];
        if (weekendDelta !== 0) {
          return weekendDelta;
        }

        return left.name.localeCompare(right.name, 'ru');
      });

    if (!eligible.length) {
      return null;
    }

    const chosen = eligible[0];
    duties[chosen.id + ':' + dayIndex] = true;

    if (day.isWeekday) {
      weekdayDutyCounts[chosen.id] += 1;
    } else {
      weekendDutyCounts[chosen.id] += 1;
    }
  }

  return duties;
}

function getRemovableWorkDays_(employee, employeeStatuses, protectedWork, coverage, days) {
  return days
    .map((day, index) => ({ day, index }))
    .filter((entry) => employeeStatuses[entry.index] === 'WORK' && !protectedWork[entry.index])
    .sort((left, right) => {
      if (left.day.isWeekday !== right.day.isWeekday) {
        return left.day.isWeekday ? 1 : -1;
      }

      const coverageDelta = coverage[right.index] - coverage[left.index];
      if (coverageDelta !== 0) {
        return coverageDelta;
      }

      return left.index - right.index;
    })
    .map((entry) => entry.index);
}

function getWeekendPairOptions_(employee, days) {
  const options = [];

  for (let index = 0; index < days.length - 1; index += 1) {
    const saturday = days[index];
    const sunday = days[index + 1];
    if (saturday.weekday !== 6 || sunday.weekday !== 0) {
      continue;
    }

    if (employee.vacationDays.has(saturday.day) || employee.vacationDays.has(sunday.day)) {
      continue;
    }

    if (employee.fixedOffWeekdays.has(saturday.weekday) || employee.fixedOffWeekdays.has(sunday.weekday)) {
      continue;
    }

    options.push({
      saturdayIndex: index,
      sundayIndex: index + 1,
    });
  }

  return options;
}

function applyScheduleToSheet_(sheet, layout, days, employees, solution, formatSources) {
  const headerNumbers = new Array(APP_CONFIG.maxDayColumns).fill('');
  const headerWeekdays = new Array(APP_CONFIG.maxDayColumns).fill('');

  for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
    headerNumbers[dayIndex] = dayIndex + 1;
    headerWeekdays[dayIndex] = APP_CONFIG.weekdayNames[days[dayIndex].weekday];
  }

  sheet.getRange(1, APP_CONFIG.dayStartColumn, 1, APP_CONFIG.maxDayColumns).setValues([headerNumbers]);
  sheet.getRange(2, APP_CONFIG.dayStartColumn, 1, APP_CONFIG.maxDayColumns).setValues([headerWeekdays]);

  layout.employeeRows.forEach((entry) => {
    const row = entry.row;
    const employee = employees.find((candidate) => candidate.id === entry.name);
    const dayValues = new Array(APP_CONFIG.maxDayColumns).fill('');

    if (!employee) {
      sheet.getRange(row, 3, 1, 2).clearContent();
      sheet.getRange(row, APP_CONFIG.dayStartColumn, 1, APP_CONFIG.maxDayColumns).clearContent();
      return;
    }

    for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
      const status = solution.statuses[employee.id][dayIndex];
      if (status === 'WORK') {
        dayValues[dayIndex] = employee.shiftHours;
      } else if (status === 'VACATION') {
        dayValues[dayIndex] = formatSources.vacationCode;
      } else {
        dayValues[dayIndex] = formatSources.offCode;
      }
    }

    sheet.getRange(row, 3).setValue(employee.targetWorkDays * employee.shiftHours);
    sheet.getRange(row, 4).setFormula(
      '=SUM(' +
        columnToLetter_(APP_CONFIG.dayStartColumn) +
        row +
        ':' +
        columnToLetter_(APP_CONFIG.dayStartColumn + APP_CONFIG.maxDayColumns - 1) +
        row +
        ')'
    );
    sheet.getRange(row, APP_CONFIG.dayStartColumn, 1, APP_CONFIG.maxDayColumns).setValues([dayValues]);

    for (let dayIndex = 0; dayIndex < days.length; dayIndex += 1) {
      const cell = sheet.getRange(row, APP_CONFIG.dayStartColumn + dayIndex);
      const status = solution.statuses[employee.id][dayIndex];
      const isDuty = Boolean(solution.duties[employee.id + ':' + dayIndex]);

      if (status === 'VACATION') {
        formatSources.samples.vacation.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      } else if (status === 'OFF') {
        formatSources.samples.off.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      } else if (isDuty) {
        formatSources.samples.duty.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      } else {
        formatSources.samples.work.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }
    }
  });

  if (days.length < APP_CONFIG.maxDayColumns) {
    const extraCols = APP_CONFIG.maxDayColumns - days.length;
    sheet.getRange(1, APP_CONFIG.dayStartColumn + days.length, 2, extraCols).clearContent();
  }
}

function detectFormatSamplesFromSheet_(sheet, layout) {
  const workCandidates = [];
  const textCandidates = {};
  const backgroundsByNumeric = {};
  const rowCount = layout.employeeRows.length;
  const range = sheet.getRange(APP_CONFIG.dataStartRow, APP_CONFIG.dayStartColumn, Math.max(rowCount, 1), APP_CONFIG.maxDayColumns);
  const displayValues = range.getDisplayValues();
  const backgrounds = range.getBackgrounds();

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let colOffset = 0; colOffset < APP_CONFIG.maxDayColumns; colOffset += 1) {
      const displayValue = String(displayValues[rowIndex][colOffset] || '').trim();
      if (!displayValue) {
        continue;
      }

      const row = APP_CONFIG.dataStartRow + rowIndex;
      const col = APP_CONFIG.dayStartColumn + colOffset;
      const background = normalizeColor_(backgrounds[rowIndex][colOffset]);
      const numeric = parseNumericCell_(displayValue);

      if (numeric !== null) {
        workCandidates.push({ row, col, background, numeric });
        backgroundsByNumeric[background] = (backgroundsByNumeric[background] || 0) + 1;
      } else {
        if (!textCandidates[displayValue]) {
          textCandidates[displayValue] = [];
        }
        textCandidates[displayValue].push({ row, col, background, value: displayValue });
      }
    }
  }

  const sortedTextEntries = Object.keys(textCandidates)
    .map((value) => ({ value, cells: textCandidates[value] }))
    .sort((left, right) => right.cells.length - left.cells.length);

  const offEntry = sortedTextEntries[0] || null;
  const vacationEntry = sortedTextEntries.find((entry) => entry.value !== (offEntry && offEntry.value)) || null;

  const numericBackgrounds = Object.keys(backgroundsByNumeric).sort((left, right) => backgroundsByNumeric[right] - backgroundsByNumeric[left]);
  const dutyBackground =
    numericBackgrounds.slice().sort((left, right) => rednessScore_(right) - rednessScore_(left))[0] || null;

  const workBackground = numericBackgrounds.find((value) => value !== dutyBackground) || dutyBackground;
  const dutyCandidate = workCandidates.find((candidate) => candidate.background === dutyBackground) || workCandidates[0];
  const workCandidate =
    workCandidates.find((candidate) => candidate.background === workBackground) ||
    workCandidates.find((candidate) => candidate.background !== dutyBackground) ||
    workCandidates[0];

  return {
    offCode: offEntry ? offEntry.value : APP_CONFIG.offCodeFallback,
    vacationCode: vacationEntry ? vacationEntry.value : APP_CONFIG.vacationCodeFallback,
    samples: {
      work: workCandidate ? sheet.getRange(workCandidate.row, workCandidate.col) : null,
      off: offEntry ? sheet.getRange(offEntry.cells[0].row, offEntry.cells[0].col) : null,
      vacation: vacationEntry ? sheet.getRange(vacationEntry.cells[0].row, vacationEntry.cells[0].col) : null,
      duty: dutyCandidate ? sheet.getRange(dutyCandidate.row, dutyCandidate.col) : null,
    },
  };
}

function refreshFormatSamples_(settingsSheet, detectedFormats) {
  const cells = getSampleCellMap_(settingsSheet);
  const values = {
    work: APP_CONFIG.defaultShiftHours,
    off: detectedFormats.offCode,
    vacation: detectedFormats.vacationCode,
    duty: APP_CONFIG.defaultShiftHours,
  };

  Object.keys(cells).forEach((key) => {
    const source = detectedFormats.samples[key];
    if (source) {
      source.copyTo(cells[key], SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      cells[key].setValue(values[key]);
    } else if (!cells[key].getDisplayValue()) {
      cells[key].setValue(values[key]);
    }
  });
}

function getFormatSampleRanges_(settingsSheet, detectedFormats) {
  const cells = getSampleCellMap_(settingsSheet);

  return {
    offCode: detectedFormats.offCode || APP_CONFIG.offCodeFallback,
    vacationCode: detectedFormats.vacationCode || APP_CONFIG.vacationCodeFallback,
    samples: {
      work: cells.work,
      off: cells.off,
      vacation: cells.vacation.getDisplayValue() ? cells.vacation : cells.off,
      duty: cells.duty.getDisplayValue() ? cells.duty : cells.work,
    },
  };
}

function getSampleCellMap_(settingsSheet) {
  return {
    work: settingsSheet.getRange(2, APP_CONFIG.sampleLegendStartColumn),
    off: settingsSheet.getRange(2, APP_CONFIG.sampleLegendStartColumn + 1),
    vacation: settingsSheet.getRange(2, APP_CONFIG.sampleLegendStartColumn + 2),
    duty: settingsSheet.getRange(2, APP_CONFIG.sampleLegendStartColumn + 3),
  };
}

function getMonthSheetLayout_(sheet) {
  const lastRow = sheet.getLastRow();
  const names = sheet.getRange(APP_CONFIG.dataStartRow, 1, Math.max(lastRow - APP_CONFIG.dataStartRow + 1, 1), 1).getDisplayValues();
  const employeeRows = [];
  let blankStreak = 0;

  for (let index = 0; index < names.length; index += 1) {
    const name = String(names[index][0] || '').trim();
    if (!name) {
      if (employeeRows.length) {
        blankStreak += 1;
        if (blankStreak >= 3) {
          break;
        }
      }
      continue;
    }

    blankStreak = 0;
    employeeRows.push({
      name,
      row: APP_CONFIG.dataStartRow + index,
    });
  }

  return {
    employeeRows,
  };
}

function normalizeDialogPayload_(payload) {
  const month = Number(payload && payload.month);
  const year = Number(payload && payload.year);
  if (!month || month < 1 || month > 12 || !year || year < 2000) {
    throw new Error('Месяц и год заданы некорректно.');
  }

  const employees = (payload.employees || [])
    .map((entry) => ({
      name: String(entry.name || '').trim(),
      normHours: Number(entry.normHours),
      vacationFrom: String(entry.vacationFrom || '').trim(),
      vacationTo: String(entry.vacationTo || '').trim(),
    }))
    .filter((entry) => entry.name);

  employees.forEach((entry) => {
    if (!entry.normHours || entry.normHours <= 0) {
      throw new Error('Для "' + entry.name + '" нужно указать положительную норму часов.');
    }
  });

  return {
    monthIndex: month - 1,
    year,
    employees,
  };
}

function normalizeNormForDialog_(value) {
  const numeric = parseNumericCell_(String(value || '').trim());
  return numeric === null ? '' : numeric;
}

function normalizeShiftHours_(value) {
  const numeric = Number(value);
  if (!numeric || numeric <= 0) {
    return APP_CONFIG.defaultShiftHours;
  }
  return numeric;
}

function buildVacationDaySet_(fromRaw, toRaw, year, monthIndex) {
  const set = new Set();
  if (!fromRaw && !toRaw) {
    return set;
  }

  if (!fromRaw || !toRaw) {
    throw new Error('Отпуск нужно задавать обеими датами: начало и конец.');
  }

  const from = parseIsoDate_(fromRaw);
  const to = parseIsoDate_(toRaw);
  if (from.getTime() > to.getTime()) {
    throw new Error('Дата начала отпуска не может быть позже даты окончания.');
  }

  const current = new Date(from.getTime());
  while (current.getTime() <= to.getTime()) {
    if (current.getFullYear() === year && current.getMonth() === monthIndex) {
      set.add(current.getDate());
    }
    current.setDate(current.getDate() + 1);
  }

  return set;
}

function parseIsoDate_(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) {
    throw new Error('Ожидалась дата в формате YYYY-MM-DD.');
  }

  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function parseFixedOffWeekdays_(value) {
  const tokens = String(value || '')
    .toLowerCase()
    .split(/[\s,;/]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);

  const map = {
    вс: 0,
    воскресенье: 0,
    пн: 1,
    понедельник: 1,
    вт: 2,
    вторник: 2,
    ср: 3,
    среда: 3,
    чт: 4,
    четверг: 4,
    пт: 5,
    пятница: 5,
    сб: 6,
    суббота: 6,
  };

  const result = new Set();
  tokens.forEach((token) => {
    if (Object.prototype.hasOwnProperty.call(map, token)) {
      result.add(map[token]);
    }
  });

  return result;
}

function buildMonthDays_(year, monthIndex) {
  const lastDay = new Date(year, monthIndex + 1, 0).getDate();
  const days = [];

  for (let day = 1; day <= lastDay; day += 1) {
    const date = new Date(year, monthIndex, day);
    const weekday = date.getDay();
    days.push({
      day,
      weekday,
      isThursday: weekday === 4,
      isWeekday: weekday >= 1 && weekday <= 5,
    });
  }

  return days;
}

function findBestTemplateSheet_(spreadsheet) {
  const parsedSheets = spreadsheet
    .getSheets()
    .map((sheet) => ({ sheet, parsed: parseMonthSheetName_(sheet.getName()) }))
    .filter((entry) => entry.parsed)
    .sort((left, right) => left.parsed.sortKey - right.parsed.sortKey);

  return parsedSheets.length ? parsedSheets[parsedSheets.length - 1].sheet : null;
}

function findTemplateSheetForTarget_(spreadsheet, year, monthIndex) {
  const targetKey = year * 12 + monthIndex;
  const parsedSheets = spreadsheet
    .getSheets()
    .map((sheet) => ({ sheet, parsed: parseMonthSheetName_(sheet.getName()) }))
    .filter((entry) => entry.parsed)
    .sort((left, right) => left.parsed.sortKey - right.parsed.sortKey);

  const previous = parsedSheets
    .filter((entry) => entry.parsed.sortKey < targetKey)
    .sort((left, right) => right.parsed.sortKey - left.parsed.sortKey)[0];

  if (previous) {
    return previous.sheet;
  }

  return parsedSheets.length ? parsedSheets[parsedSheets.length - 1].sheet : null;
}

function getSuggestedTargetMonth_(sheetName) {
  const parsed = parseMonthSheetName_(sheetName);
  if (!parsed) {
    const now = new Date();
    return { year: now.getFullYear(), month: now.getMonth() };
  }

  const nextMonth = parsed.month + 1;
  if (nextMonth > 11) {
    return { year: parsed.year + 1, month: 0 };
  }

  return { year: parsed.year, month: nextMonth };
}

function parseMonthSheetName_(sheetName) {
  const match = /^(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)\s+(\d{4})(?:\s+\(\d+\))?$/i.exec(
    String(sheetName || '').trim()
  );
  if (!match) {
    return null;
  }

  const month = APP_CONFIG.monthNames.indexOf(match[1].toLowerCase());
  if (month === -1) {
    return null;
  }

  const year = Number(match[2]);
  return {
    month,
    year,
    sortKey: year * 12 + month,
  };
}

function buildMonthSheetName_(year, monthIndex) {
  return APP_CONFIG.monthNames[monthIndex] + ' ' + year;
}

function copySheetWithUniqueName_(spreadsheet, templateSheet, baseName) {
  const copy = templateSheet.copyTo(spreadsheet);
  let candidate = baseName;
  let suffix = 2;

  while (spreadsheet.getSheetByName(candidate)) {
    candidate = baseName + ' (' + suffix + ')';
    suffix += 1;
  }

  copy.setName(candidate);
  spreadsheet.setActiveSheet(copy);
  return copy;
}

function asBoolean_(value, fallback) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return fallback;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  const normalized = String(value).trim().toLowerCase();
  if (['true', '1', 'yes', 'y', 'да', 'истина'].indexOf(normalized) !== -1) {
    return true;
  }
  if (['false', '0', 'no', 'n', 'нет', 'ложь'].indexOf(normalized) !== -1) {
    return false;
  }

  return fallback;
}

function parseNumericCell_(value) {
  const normalized = String(value || '')
    .trim()
    .replace(',', '.');
  if (!normalized || !/^-?\d+(\.\d+)?$/.test(normalized)) {
    return null;
  }
  return Number(normalized);
}

function setStatus_(statusRow, dayIndex, nextStatus) {
  const current = statusRow[dayIndex];
  if (!current || current === nextStatus) {
    statusRow[dayIndex] = nextStatus;
    return true;
  }

  return false;
}

function normalizeColor_(color) {
  if (!color) {
    return '#ffffff';
  }

  const normalized = String(color).trim().toLowerCase();
  if (normalized.charAt(0) === '#') {
    return normalized;
  }

  if (normalized.length === 6) {
    return '#' + normalized;
  }

  return normalized;
}

function rednessScore_(color) {
  const normalized = normalizeColor_(color).replace('#', '');
  if (normalized.length !== 6) {
    return -999;
  }

  const red = parseInt(normalized.slice(0, 2), 16);
  const green = parseInt(normalized.slice(2, 4), 16);
  const blue = parseInt(normalized.slice(4, 6), 16);
  return red - Math.max(green, blue);
}

function columnToLetter_(column) {
  let current = column;
  let result = '';
  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result;
}
