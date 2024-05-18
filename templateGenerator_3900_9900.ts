////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////// INTERFACES & ENUMS ////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////

interface Combinations {
  [key: string]: {
    'in-person': number;
    'online': number;
  }
}

interface TeachingModeRows {
  'in-person': number[];
  'online': number[];
}

enum COLORS {
  SKY_BLUE = 'C9DAF8',
  CYAN = 'B5E3E8',
  CORAL = 'E06666',
  DARK_RED = 'C00000',
  RED = 'FF0000',
  SALMON = 'FBDAD7',
  MISTY_ROSE = 'F4CCCC',
  PINK = 'FFC7CE',
  THISTLE = 'D9D2E9',
  SILVER = 'B7B7B7',
  GREEN = 'AEDE7A',
  GOLD = 'F2CD5E',
  YELLOW = 'FFE599',
  LEMON = 'FFF2CC',
  WHITE = 'FFFFFF',
  BLACK = '000000',
}

////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////// HELPER FUNCTIONS /////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * Converts a 24-hour time string to a 12-hour format string
 * @param time24 The 24-hour time string (e.g., '14:00')
 * @returns The 12-hour format string (e.g., '2pm')
 */
function convertTo12HourFormat(time24: string): string {
  // Split the time string into hours and minutes
  const [hours] = time24.split(':').map(arg => Number(arg));

  // Construct and return the 12-hour time string
  return `${hours % 12 || 12}${hours >= 12 ? 'pm' : 'am'}`;
}

/**
 * Parses a 12-hour format time string and returns the hour in 24-hour format
 * @param time The 12-hour format time string (e.g., '2pm')
 * @returns The hour in 24-hour format (e.g., 14)
 */
function parseHour(time: string): number {
  const hourStr = time.slice(0, -2); // Remove the 'am' or 'pm' part
  const hour = Number(hourStr);
  if ((time.endsWith('pm') && hour !== 12) || (time.endsWith('am') && hour === 12)) {
    return (hour + 12) % 24;
  }
  return hour;
}

/**
 * Processes class schedules from a given Excel worksheet
 * @param timetable The Excel worksheet containing the timetable data
 * @param startRow The starting row of the timetable data
 * @param endRow The ending row of the timetable data
 * @param classTimesCol The column containing class times
 * @param classLocationsCol The column containing class locations
 * @returns An object containing the class schedule combinations
 */
function processClassSchedules(timetable: ExcelScript.Worksheet, startRow: number, endRow: number, classTimesCol: string, classLocationsCol: string): Combinations {
  const classTimes = timetable.getRange(`${classTimesCol}${startRow}:${classTimesCol}${endRow}`).getValues().map((row: string[]) => row[0].toString());
  const classLocations = timetable.getRange(`${classLocationsCol}${startRow}:${classLocationsCol}${endRow}`).getValues().map((row: string[]) => row[0].toString());

  const combinations: Combinations = {};

  for (let i = 0; i < classTimes.length; i++) {
    const [day, begin, , end] = classTimes[i].split(' ');
    const isOnline = classLocations[i] === 'Online (ONLINE)';

    const key = `${day} ${convertTo12HourFormat(begin)} - ${convertTo12HourFormat(end)}`;
    const { 'in-person': inPersonCount, 'online': onlineCount } = combinations[key] || { 'in-person': 0, 'online': 0 };

    combinations[key] = {
      'in-person': isOnline ? inPersonCount : inPersonCount + 1,
      'online': isOnline ? onlineCount + 1 : onlineCount
    };
  }

  return combinations;
}

/**
 * Orders the class schedule combinations by weekday and start time
 * @param combinations The class schedule combinations
 * @returns The ordered class schedule combinations
 */
function orderCombinations(combinations: Combinations): Combinations {
  const weekdaysOrder = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
  const orderedCombinations: Combinations = {};

  Object.keys(combinations)
    .sort((a, b) => {
      const [dayA, timeA] = a.split(' ');
      const [dayB, timeB] = b.split(' ');

      // Compare by day
      if (dayA !== dayB) {
        return weekdaysOrder.indexOf(dayA) - weekdaysOrder.indexOf(dayB);
      } else {
        // If days are equal, compare by start time
        return parseHour(timeA) - parseHour(timeB);
      }
    })
    .forEach(key => {
      orderedCombinations[key] = combinations[key];
    });

  return orderedCombinations;
}

/**
 * Aligns the text in the given range to be centered horizontally and vertically
 * @param range The Excel range to align
 */
function alignCenterMiddle(range: ExcelScript.Range) {
  range.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}

/**
 * Sets a conditional formatting rule for blank cells in the given range
 * @param dataRange The Excel range to apply the rule to
 * @param color The fill color for the blank cells
 */
function setBlankRule(dataRange: ExcelScript.Range, color: string) {
  const blankRule: ExcelScript.ConditionalPresetCriteriaRule = {
    criterion: ExcelScript.ConditionalFormatPresetCriterion.blanks
  };

  const conditionalFormat = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.presetCriteria);
  const presetFormat = conditionalFormat.getPreset();
  presetFormat.getFormat().getFill().setColor(color);
  presetFormat.setRule(blankRule);
}

/**
 * Adds borders to the specified range in an Excel worksheet
 * @param range The Excel range to add borders to
 * @param insideHorizontal If true, adds inside horizontal borders
 * @param insideVertical If true, adds inside vertical borders
 * @param weight The weight of the border lines (thin or medium)
 */
function addBorders(range: ExcelScript.Range, insideHorizontal: Boolean, insideVertical: Boolean, weight: ExcelScript.BorderWeight) {
  const borders = [
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom
  ];

  if (insideHorizontal) {
    borders.push(ExcelScript.BorderIndex.insideHorizontal);
  }

  if (insideVertical) {
    borders.push(ExcelScript.BorderIndex.insideVertical);
  }

  for (const border of borders) {
    range.getFormat().getRangeBorder(border).setStyle(ExcelScript.BorderLineStyle.continuous);
    range.getFormat().getRangeBorder(border).setColor(COLORS['BLACK']);
    range.getFormat().getRangeBorder(border).setWeight(weight);
  }
}

/**
 * Creates the left-most columns of the worksheet for tutor names and their zIDs
 * @param sheet The Excel worksheet to update
 */
function createLeftMostMostColumns(sheet: ExcelScript.Worksheet) {
  let dataRange = sheet.getRange('A4:B4');
  dataRange.getFormat().getFill().setColor(COLORS['BLACK']);
  dataRange.getFormat().getFont().setColor(COLORS['WHITE']);
  dataRange.setValues([['Name', 'zID']]);
  dataRange.getFormat().getFont().setBold(true);
  alignCenterMiddle(dataRange);

  sheet.getRange('A5:A34').getFormat().setColumnWidth(130);
  sheet.getRange('B5:B34').getFormat().setColumnWidth(90);

  dataRange = sheet.getRange('A5:B34');
  dataRange.getFormat().getFill().setColor(COLORS['SKY_BLUE']);
  dataRange.getFormat().getFont().setBold(true);
  dataRange.setValues(Array(30).fill(Array(2).fill('[fill in here]')));

  addBorders(dataRange, true, true, ExcelScript.BorderWeight.thin);
}

/**
 * Records the class preferences in the worksheet
 * @param sheet The Excel worksheet to update
 * @param classes The array of class descriptions
 * @param types The array of class types (e.g., 'In-person' or 'Online')
 * @param counts The array of class counts
 * @returns The character code of the last column used
 */
function recordPreferences(sheet: ExcelScript.Worksheet, classes: string[], types: string[], counts: number[]) {
  let endCol = String.fromCharCode(64 + classes.length);
  let dataRange = sheet.getRange(`A1:${endCol}3`);
  dataRange.setValues([classes, types, counts]);
  dataRange.getFormat().getFont().setBold(true);
  dataRange.getFormat().autofitColumns();
  dataRange.getFormat().autofitRows();
  dataRange.getFormat().getFill().setColor(COLORS['THISTLE']);
  alignCenterMiddle(dataRange);

  sheet.getRange(`A2:${endCol}2`).getFormat().getFill().setColor(COLORS['SILVER']);

  sheet.getRange('A:A').insert(ExcelScript.InsertShiftDirection.right);
  sheet.getRange('A:A').insert(ExcelScript.InsertShiftDirection.right);
  let endColCharCode = 64 + classes.length + 2;
  endCol = String.fromCharCode(endColCharCode);

  dataRange = sheet.getRange(`C4:${endCol}4`);
  setValueAndColor(dataRange, 'Preference', COLORS['YELLOW']);

  setBlankRule(sheet.getRange(`C5:${endCol}34`), COLORS['MISTY_ROSE']);

  createLeftMostMostColumns(sheet);

  addBorders(sheet.getRange(`C1:${endCol}34`), true, true, ExcelScript.BorderWeight.thin);

  return endColCharCode;
}

/**
 * Creates the right-most columns of the worksheet for additional class details
 * @param sheet The Excel worksheet to update
 * @param colCharCode The character code of the last column used
 */
function createRightMostColumns(sheet: ExcelScript.Worksheet, colCharCode: number) {
  let colCode = colCharCode;

  const titles = ['Done', 'Ideal # classes', 'Max # classes', 'Notes'];
  for (const title of titles) {
    colCode += 1;
    const col = String.fromCharCode(colCode);
    let dataRange = sheet.getRange(`${col}1:${col}4`);
    dataRange.merge();
    dataRange.getFormat().getFont().setBold(true);
    dataRange.getFormat().setColumnWidth(85);
    setValueAndColor(dataRange, title, COLORS['LEMON']);

    dataRange = sheet.getRange(`${col}5:${col}34`);
    if (title === 'Done') {
      dataRange.clear(ExcelScript.ClearApplyTo.contents);
      const dataValidation = dataRange.getDataValidation();
      dataValidation.setIgnoreBlanks(true);
      const validationCriteria: ExcelScript.ListDataValidation = {
        inCellDropDown: true,
        source: 'TODO,Yes'
      };
      const validationRule: ExcelScript.DataValidationRule = {
        list: validationCriteria
      };
      dataValidation.setRule(validationRule);
      setValueAndColor(dataRange, 'TODO', COLORS['GREEN']);

      const todoValRule: ExcelScript.ConditionalTextComparisonRule = {
        operator: ExcelScript.ConditionalTextOperator.contains,
        text: 'TODO'
      };

      const textConditionFormat = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.containsText).getTextComparison();
      textConditionFormat.getFormat().getFill().setColor(COLORS['CORAL']);
      textConditionFormat.setRule(todoValRule);

      setBlankRule(dataRange, COLORS['GOLD']);
    } else {
      setBlankRule(dataRange, COLORS['PINK']);

      const zeroValRule: ExcelScript.ConditionalCellValueRule = {
        formula1: '0',
        operator: ExcelScript.ConditionalCellValueOperator.equalTo
      };

      const cellValConditionalFormat = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue).getCellValue();
      cellValConditionalFormat.getFormat().getFill().setColor(COLORS['PINK']);
      cellValConditionalFormat.getFormat().getFont().setColor(COLORS['DARK_RED']);
      cellValConditionalFormat.setRule(zeroValRule);
    }
  }

  addBorders(sheet.getRange(`${String.fromCharCode(colCharCode)}1:${String.fromCharCode(colCharCode + 3)}34`), true, false, ExcelScript.BorderWeight.thin);
  addBorders(sheet.getRange(`${String.fromCharCode(colCharCode + 4)}1:${String.fromCharCode(colCharCode + 4)}34`), true, false, ExcelScript.BorderWeight.thin);
}

/**
 * Finds the maximum number of in-person and online classes in the schedule
 * @param schedule The class schedule combinations
 * @returns A tuple containing the maximum number of in-person and online classes
 */
function maxValues(schedule: Combinations): [number, number] {
  const inPersonValues = Object.values(schedule).map(slot => slot['in-person']);
  const onlineValues = Object.values(schedule).map(slot => slot['online']);
  const maxInPerson = Math.max(...inPersonValues);
  const maxOnline = Math.max(...onlineValues);
  return [maxInPerson, maxOnline];
}

/**
 * Creates rows for in-person or online classes in the worksheet for the footer.
 * @param sheet The Excel worksheet
 * @param teachingMode The teaching mode ('In-person' or 'Online')
 * @param currRow The current row in the worksheet
 * @param endCol The ending column for the range
 */
function createTimetableRows(sheet: ExcelScript.Worksheet, teachingMode: string, currRow: number, endCol: string) {
  let dataRange = sheet.getRange(`C${currRow}:${endCol}${currRow + 1}`);
  setValueAndColor(dataRange, '-', COLORS['SALMON']);

  const emptyCell = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom).getCustom();
  emptyCell.getFormat().getFill().setColor(COLORS['RED']);
  emptyCell.getRule().setFormula(`=C${currRow}=""`);

  const extraCell = dataRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom).getCustom();
  extraCell.getFormat().getFill().setColor(COLORS['CYAN']);
  extraCell.getRule().setFormula(`=C${currRow}="-"`);

  dataRange = sheet.getRange(`B${currRow}:B${currRow + 1}`);
  dataRange.merge();
  dataRange.setValue(teachingMode);
  alignCenterMiddle(dataRange);
}

/**
 * Clears the contents of the specified cell range
 * @param sheet The Excel worksheet to update
 * @param cellRange The range of cells to clear
 */
function clearCellContents(sheet: ExcelScript.Worksheet, col: string, teachingModeCount: number, rows: number[]) {
  for (let i = 0; i != teachingModeCount; i++) {
    const dataRange = sheet.getRange(`${col}${rows[i]}:${col}${rows[i] + 1}`);
    dataRange.setValue('');
  }
}

/**
 * Creates a timetable in the specified worksheet
 * @param sheet The Excel worksheet to update
 * @param startRow The starting row of the timetable
 * @param endRow The ending row of the timetable
 * @param columns The columns to use for the timetable
 * @param title The title of the timetable
 */
function createTimetable(sheet: ExcelScript.Worksheet, schedule: Combinations, classes: string[], types: string[], counts: number[]) {
  const [maxInPerson, maxOnline] = maxValues(schedule);

  let currRow = 35;
  const endCol = String.fromCharCode(64 + classes.length + 2);
  const dataRange = sheet.getRange(`C${currRow}:${endCol}${currRow}`);
  dataRange.setValues([classes]);
  dataRange.getFormat().getFill().setColor(COLORS['THISTLE']);
  dataRange.getFormat().getFont().setBold(true);
  alignCenterMiddle(dataRange);
  const height = dataRange.getFormat().getRowHeight();
  dataRange.getFormat().setRowHeight(height * 2);

  currRow += 1;

  const teachingModeRows: TeachingModeRows = {
    'in-person': [],
    'online': []
  }

  for (let i = 0; i < maxInPerson; i++) {
    teachingModeRows['in-person'].push(currRow);
    createTimetableRows(sheet, 'In-person', currRow, endCol);
    currRow += 2;
  }

  for (let i = 0; i < maxOnline; i++) {
    teachingModeRows['online'].push(currRow);
    createTimetableRows(sheet, 'Online', currRow, endCol);
    currRow += 2;
  }

  let col = 66;
  for (let i = 0; i < classes.length; i++) {
    col++;
    if (types[i].toLocaleLowerCase() === 'in-person') {
      clearCellContents(sheet, String.fromCharCode(col), counts[i], teachingModeRows['in-person']);
    } else {
      clearCellContents(sheet, String.fromCharCode(col), counts[i], teachingModeRows['online']);
    }
  }
}

/**
 * Sets the value and fill color of the specified range, and aligns the text to the center
 * @param range The Excel range to update
 * @param val The value to set
 * @param color The fill color to set
 */
function setValueAndColor(range: ExcelScript.Range, val: string, color: string) {
  alignCenterMiddle(range);
  range.getFormat().getFill().setColor(color);
  range.setValue(val);
}

/**
 * Adds an instruction block to the worksheet with a header and key-value pairs
 * @param sheet The Excel worksheet to update
 * @param header The header text for the instruction block
 * @param startCol The starting column number for the block
 * @param endCol The ending column number for the block
 * @param color The fill color for the block
 * @param obj The object containing key-value pairs to display
 */
function addInstructionBlock(sheet: ExcelScript.Worksheet, header: string, startCol: number, endCol: number, color: string, obj: Object, ) {
  let row = 1;
  let dataRange = sheet.getRange(`${String.fromCharCode(startCol)}${row}:${String.fromCharCode(endCol)}${row}`);
  dataRange.merge();
  dataRange.getFormat().getFont().setBold(true);
  setValueAndColor(dataRange, header, color);

  for (const [key, value] of Object.entries(obj)) {
    row += 1;
    dataRange = sheet.getRange(`${String.fromCharCode(startCol)}${row}:${String.fromCharCode(startCol)}${row}`);
    setValueAndColor(dataRange, key, color);

    dataRange = sheet.getRange(`${String.fromCharCode(startCol + 1)}${row}:${String.fromCharCode(endCol)}${row}`);
    dataRange.merge();
    setValueAndColor(dataRange, value, color);
  }
}

/**
 * Adds instruction rows at the top of the worksheet
 * @param sheet The Excel worksheet to update
 * @param classesLen The number of classes to calculate the instruction block positions
 */
function addInstructionRows(sheet: ExcelScript.Worksheet, classesLen: number) {
  sheet.getRange('1:1').insert(ExcelScript.InsertShiftDirection.down);
  sheet.getRange('1:1').insert(ExcelScript.InsertShiftDirection.down);
  sheet.getRange('1:1').insert(ExcelScript.InsertShiftDirection.down);
  sheet.getRange('1:1').insert(ExcelScript.InsertShiftDirection.down);

  const preferences = {
    '1': 'Preferred',
    '2': 'Possible'
  }

  let col = 63 + classesLen / 2;

  addInstructionBlock(sheet, 'Preference', col, col + 2, COLORS['YELLOW'], preferences);
}

/**
 * Updates the worksheet with class schedule data and formatting
 * @param sheet The Excel worksheet to update
 * @param schedule The class schedule combinations
 * @returns The range that was updated
 */
function updateWorksheet(sheet: ExcelScript.Worksheet, schedule: Combinations) {
  const classes: string[] = [];
  const types: string[] = [];
  const counts: number[] = [];

  for (const [key, value] of Object.entries(schedule)) {
    if (value['in-person']) {
      classes.push(key);
      types.push('In-person');
      counts.push(value['in-person']);
    }

    if (value['online']) {
      classes.push(key);
      types.push('Online');
      counts.push(value['online']);
    }
  }

  if (classes.length !== types.length || classes.length !== counts.length || types.length !== counts.length) {
    throw new Error('Something went wrong!');
  }

  createRightMostColumns(sheet, recordPreferences(sheet, classes, types, counts));

  createTimetable(sheet, schedule, classes, types, counts);

  let col = 67;
  for (let i = 0; i < classes.length; i++) {
    const dataRange = sheet.getRange(`${String.fromCharCode(col)}1:${String.fromCharCode(col)}34`);
    addBorders(dataRange, false, false, ExcelScript.BorderWeight.medium);
    col++;
  }

  col = 67;
  const dataRange = sheet.getRange(`${String.fromCharCode(col)}35:${String.fromCharCode(col + classes.length - 1)}35`);
  addBorders(dataRange, false, true, ExcelScript.BorderWeight.medium);

  addInstructionRows(sheet, classes.length);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////// MAIN FUNCTIONS //////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * Main function to generate the class schedule worksheet
 * @param workbook The Excel workbook object
 * @param course The course code (either 'COMP3900' or 'COMP9900')
 * @param startRow The starting row of the timetable data
 * @param endRow The ending row of the timetable data
 * @param classTimesCol The column containing class times
 * @param classLocationsCol The column containing class locations
 */
function main(workbook: ExcelScript.Workbook, course: 'COMP3900' | 'COMP9900', startRow: number = 31, endRow: number = 50, classTimesCol: string = 'A', classLocationsCol: string = 'B') {
  const timetable = workbook.getWorksheet('TT');
  if (!timetable) {
    throw new Error('Timetable not found');
  }

  if (startRow > endRow) {
    throw new Error('Start row cannot be after End row');
  }

  if (workbook.getWorksheet(course)) {
    throw new Error(`Sheet ${course} already exists`);
  }

  const worksheet = workbook.addWorksheet(course);

  const schedule = processClassSchedules(timetable, startRow, endRow, classTimesCol, classLocationsCol);
  updateWorksheet(worksheet, orderCombinations(schedule));
  worksheet.getFreezePanes().freezeRows(8);
}
