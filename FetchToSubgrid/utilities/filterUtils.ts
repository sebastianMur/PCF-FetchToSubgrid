import { IColumn, IComboBoxOption } from '@fluentui/react';
import { AttributeType } from '../@types/enums';
import { Dispatch, SetStateAction } from 'react';

export enum Condition {
  'eq' = 0,
  'ne' = 1,
  'gt' = 2, // GreaterThan
  'lt' = 3, // LessThan
  'ge' = 4, // GreaterEqual
  'le' = 5, // LessEqual
  'like' = 6,
  'null' = 12,
  'not-null' = 13,
  'yesterday' = 14,
  'today' = 15,
  'tomorrow' = 16,
  'last-seven-days' = 17,
  'next-seven-days' = 18,
  'last-week' = 19,
  'this-week' = 20,
  'next-week' = 21,
  'last-month' = 22,
  'this-month' = 23,
  'next-month' = 24,
  'on' = 25,
  'on-or-before' = 26,
  'on-or-after' = 27,
  'last-year' = 28,
  'this-year' = 29,
  'next-year' = 30,
  'last-x-hours' = 31,
  'next-x-hours' = 32,
  'last-x-days' = 33,
  'next-x-days' = 34,
  'last-x-weeks' = 35,
  'next-x-weeks' = 36,
  'last-x-months' = 37,
  'next-x-months' = 38,
  'last-x-years' = 39,
  'next-x-years' = 40,
  'contains' = 49,
  'does-not-contain' = 50,
  'user-language' = 51,
  'olderthan-x-months' = 53,
  'begins-with' = 54,
  'does-not-begin-with' = 55,
  'ends-with' = 56,
  'does-not-end-with' = 57,
  'this-fiscal-year' = 58,
  'this-fiscal-period' = 59,
  'next-fiscal-year' = 60,
  'next-fiscal-period' = 61,
  'last-fiscal-year' = 62,
  'last-fiscal-period' = 63,
  'last-x-fiscal-years' = 64,
  'last-x-fiscal-periods' = 65,
  'next-x-fiscal-years' = 66,
  'next-x-fiscal-periods' = 67,
  'in-fiscal-year' = 68,
  'in-fiscal-period' = 69,
  'in-fiscal-period-and-year' = 70,
  'in-or-before-fiscal-period-and-year' = 71,
  'in-or-after-fiscal-period-and-year' = 72,
  'olderthan-x-years' = 82,
  'olderthan-x-weeks' = 83,
  'olderthan-x-days' = 84,
  'olderthan-x-hours' = 85,
  'olderthan-x-minutes' = 86,
  'contain-values' = 87,
  'not-contain-values' = 88,
}

const findConditionName = (condition: number) => {
  const indexOfCondition = Object.values(Condition).indexOf(condition);
  return Object.keys(Condition)[indexOfCondition];
};

export const textOptionList = [
  { text: 'Equals', key: 0 },
  { text: 'Does not equal', key: 1 },
  { text: 'Contains', key: 49 },
  { text: 'Does not contain', key: 50 },
  { text: 'Begins with', key: 54 },
  { text: 'Does not begin with', key: 55 },
  { text: 'Ends with', key: 56 },
  { text: 'Does not end with', key: 57 },
  { text: 'Contains data', key: 13 },
  { text: 'Does not contain data', key: 12 },
];

export const multiSelectOptionList = [
  { text: 'Equals', key: 0 },
  { text: 'Does not equal', key: 1 },
  { text: 'Contains values', key: 87 },
  { text: 'Does not contain values', key: 88 },
  { text: 'Contains data', key: 13 },
  { text: 'Does not contain data', key: 12 },
];

export const twoOptionList = [
  { text: 'Equals', key: 0 },
  { text: 'Contains data', key: 13 },
  { text: 'Does not contain data', key: 12 },
];

export const numberOptionList = [
  { text: 'Equals', key: 0 },
  { text: 'Does not equal', key: 1 },
  { text: 'Contains data', key: 13 },
  { text: 'Does not contain data', key: 12 },
  { text: 'Greater than', key: 2 },
  { text: 'Greater than or equal to', key: 4 },
  { text: 'Less than', key:  3 },
  { text: 'Less than or equal to', key: 5 },
];

export const languageOptionList = [
  { text: 'Equals', key: 0 },
  { text: 'Does not equal', key: 1 },
  { text: 'Contains data', key: 13 },
  { text: 'Does not contain data', key: 12 },
  { text: 'Equals current user language', key: 51 },
];

export const dateOptionList = [
  { text: 'On', key: 25 },
  { text: 'On or before', key: 26 },
  { text: 'On or after', key: 27 },
  { text: 'Today', key: 15 },
  { text: 'Yesterday', key: 14 },
  { text: 'Tomorrow', key: 16 },
  { text: 'This week', key: 20 },
  { text: 'This month', key: 23 },
  { text: 'This year', key: 29 },
  { text: 'This fiscal period', key: 59 },
  { text: 'This fiscal year', key: 58 },
  { text: 'Next week', key: 21 },
  { text: 'Next 7 days', key: 18 },
  { text: 'Next month', key: 24 },
  { text: 'Next year', key: 30 },
  { text: 'Next fiscal period', key: 61 },
  { text: 'Next fiscal year', key: 60 },
  { text: 'Next X hours', key: 32 },
  { text: 'Next X days', key: 34 },
  { text: 'Next X weeks', key: 36 },
  { text: 'Next X months', key: 38 },
  { text: 'Next X years', key: 40	},
  { text: 'Next X fiscal periods', key: 67 },
  { text: 'Next X fiscal years', key: 66 },
  { text: 'Last week', key: 19 },
  { text: 'Last 7 days', key: 17 },
  { text: 'Last month', key: 22 },
  { text: 'Last year', key: 28 },
  { text: 'Last fiscal period', key: 63 },
  { text: 'Last fiscal year', key: 62 },
  { text: 'Last X hours', key: 31 },
  { text: 'Last X days', key: 33 },
  { text: 'Last X weeks', key: 35 },
  { text: 'Last X months', key: 37 },
  { text: 'Last X years', key: 39 },
  { text: 'Last X fiscal periods', key: 65 },
  { text: 'Last X fiscal years', key: 64 }, // Value = 0x40
  { text: 'Older than X minutes', key: 86 },
  { text: 'Older than X hours', key: 85 },
  { text: 'Older than X days', key: 84 },
  { text: 'Older than X weeks', key: 83 },
  { text: 'Older than X months', key: 53 },
  { text: 'Older than X years', key: 82 },
  { text: 'In fiscal year', key: 68 }, // input is optionset
  { text: 'In fiscal period', key: 69 }, // input is optionset
  { text: 'In fiscal period and year', key: 70 }, // input is optionset
  { text: 'In or after fiscal period and year', key: 72 }, // input is optionset
  { text: 'In or before fiscal period and year', key: 71 }, // input is optionset
  { text: 'Contains data (any time)', key: 13 },
  { text: 'Does not contain data', key: 12 },
];

export const fiscalPeriodOptions: IComboBoxOption[] = [
  { text: 'Quarter 1', key: 1 },
  { text: 'Quarter 2', key: 2 },
  { text: 'Quarter 3', key: 3 },
  { text: 'Quarter 4', key: 4 },
];

export const durationList: IComboBoxOption[] = [
  { text: '1 minute', key: '1' },
  { text: '15 minutes', key: '15' },
  { text: '30 minutes', key: '30' },
  { text: '45 minutes', key: '45' },
  { text: '1 hour', key: '60' },
  { text: '1.5 hours', key: '90' },
  { text: '2 hours', key: '120' },
  { text: '2.5 hours', key: '150' },
  { text: '3 hours', key: '180' },
  { text: '3.5 hours', key: '210' },
  { text: '4 hours', key: '240' },
  { text: '4.5 hours', key: '270' },
  { text: '5 hours', key: '300' },
  { text: '5.5 hours', key: '330' },
  { text: '6 hours', key: '360' },
  { text: '6.5 hours', key: '390' },
  { text: '7 hours', key: '420' },
  { text: '7.5 hours', key: '450' },
  { text: '8 hours', key: '480' },
  { text: '1 day', key: '1440' },
  { text: '2 days', key: '2880' },
  { text: '3 days', key: '4320' },
];

export const getFilterItems = (
  selectedColumn: IColumn,
  handleRemoveFilter: () => void,
  handleSortItems: (sortType: string) => void,
  setIsFilterPopupVisible: Dispatch<SetStateAction<boolean>>) => {
  const items = [
    {
      key: 'sortAsc',
      text: 'A to Z',
      canCheck: true,
      isChecked: selectedColumn.isSorted ? selectedColumn.isSortedDescending === false : false,
      iconProps: { iconName: 'SortUp' },
      onClick: () => handleSortItems('ascending'),
    },
    {
      key: 'sortDesc',
      text: 'Z to A',
      canCheck: true,
      isChecked: selectedColumn.isSorted ? selectedColumn.isSortedDescending === true : false,
      iconProps: { iconName: 'SortDown' },
      onClick: () => handleSortItems('descending'),
    },
    {
      key: 'filterBy',
      text: 'Filter By',
      iconProps: { iconName: 'Filter' },
      onClick: () => setIsFilterPopupVisible(true),
    },
  ];

  if (selectedColumn.isFiltered) {
    items.push({
      key: 'clearFilter',
      text: 'Clear Filter',
      iconProps: { iconName: 'ClearFilter' },
      onClick: handleRemoveFilter,
    });
  }

  return items;
};

export const getOptionsByAttributeType = (attributeType: number) => {
  switch (attributeType) {
    case AttributeType.DateTime:
      return dateOptionList;

    case AttributeType.Number:
    case AttributeType.Money:
    case AttributeType.DecimalNumber:
    case AttributeType.FloatNumber:
      return numberOptionList;

    case AttributeType.MultiselectPickList:
      return multiSelectOptionList;

    case AttributeType.TwoOptions:
    case AttributeType.Status:
      return twoOptionList;

    default:
      return textOptionList;
  }
};

export const fiscalYearOptions = () => {
  let fiscalYear = 1970;
  const options = [];
  while (fiscalYear <= new Date().getFullYear() + 15) {
    options.push({ text: `FY${fiscalYear}`, key: fiscalYear });
    fiscalYear++;
  }
  return options as IComboBoxOption[];
};

export const addFilterIconToColumn = (fieldName: string, columns: IColumn[]) => {
  const column = columns.find(column => column.fieldName === fieldName);
  if (column) {
    column.isFiltered = true;
  }
};

export const removeFilterIconToColumn = (fieldName: string, columns: IColumn[]) => {
  const column = columns.find(column => column.fieldName === fieldName);
  if (column) {
    column.isFiltered = false;
  }
};

export const getFilterOptions = (column: IColumn | undefined) => {
  if (column === undefined) return [];

  switch (column.data) {
    case 'Whole.None':
    case 'Decimal':
    case 'FP':
    case 'Currency':
    case 'Whole.Duration':
    case 'Whole.TimeZone':
      return numberOptionList;

    case 'MultiSelectPicklist':
      return multiSelectOptionList;

    case 'Whole.Language':
      return languageOptionList;

    case 'DateAndTime.DateOnly':
    case 'DateAndTime.DateAndTime':
      return dateOptionList;

    case 'SingleLine.Text':
    case 'Multiple':
    default:
      return textOptionList;
  }
};

export const formatCondition = (
  option: number | string,
  value: string,
  attributeFormat: string) => {
  let currentValue = value;
  if (attributeFormat === '2' || attributeFormat === '3') {
    currentValue = value.replace(/,|\//i, '.');
  }

  switch (option) {
    case Condition.contains:
      return {
        conditionOperator: 'like',
        value: `%${currentValue}%`,
      };

    case Condition['begins-with']:
      return {
        conditionOperator: 'like',
        value: `${currentValue}%`,
      };

    case Condition['ends-with']:
      return {
        conditionOperator: 'like',
        value: `%${currentValue}`,
      };

    case Condition['does-not-end-with']:
      return {
        conditionOperator: 'not-like',
        value: `%${currentValue}`,
      };

    case Condition['does-not-contain']:
      return {
        conditionOperator: 'not-like',
        value: `%${currentValue}%`,
      };

    case Condition['does-not-begin-with']:
      return {
        conditionOperator: 'not-like',
        value: `${currentValue}%`,
      };

    case Condition['user-language']:
      return {
        conditionOperator: 'eq',
        value: currentValue,
      };
  }

  return {
    value: currentValue,
    conditionOperator: findConditionName(option as number),
  };
};

const isFiscalPeriodAndYear = (option: number) =>
  [Condition['in-fiscal-period-and-year'], Condition['in-or-after-fiscal-period-and-year'],
    Condition['in-or-before-fiscal-period-and-year']].includes(option);

const isMultipleInput = (option: number, columnType: string) =>
  ['OptionSet', 'TwoOptions', 'MultiSelectPicklist', 'Lookup.Simple',
    'Whole.TimeZone'].includes(columnType) && [Condition.eq, Condition.ne,
    Condition['contain-values'], Condition['not-contain-values']].includes(option);

export const isLookupOrDropdown = (columnType: string) =>
  ['OptionSet', 'TwoOptions', 'MultiSelectPicklist', 'Lookup.Simple'].includes(columnType);

export const getConditionByNumber = (number: number): string | undefined => {
  for (const key in Condition) {
    if (Condition[key as keyof typeof Condition] === number) {
      return key;
    }
  }
  return undefined;
};

export const isInputHidden = (option: number) =>
  [Condition['not-null'], Condition.null, Condition.today, Condition.yesterday,
    Condition.tomorrow, Condition['this-week'], Condition['this-month'], Condition['this-year'],
    Condition['this-fiscal-period'], Condition['this-fiscal-year'], Condition['next-week'],
    Condition['next-seven-days'], Condition['next-month'], Condition['next-fiscal-period'],
    Condition['next-year'], Condition['next-fiscal-year'], Condition['last-week'],
    Condition['last-month'], Condition['last-year'], Condition['last-fiscal-year'],
    Condition['last-seven-days'], Condition['last-fiscal-period'],
    Condition['user-language']].includes(option);

export const isDateTextInput = (option: number) =>
  [Condition['next-x-hours'], Condition['next-x-days'], Condition['next-x-weeks'],
    Condition['next-x-months'], Condition['next-x-years'], Condition['next-x-fiscal-periods'],
    Condition['next-x-fiscal-years'], Condition['last-x-hours'], Condition['last-x-days'],
    Condition['last-x-weeks'], Condition['last-x-months'], Condition['last-x-years'],
    Condition['last-x-fiscal-periods'], Condition['last-x-fiscal-years'],
    Condition['olderthan-x-minutes'], Condition['olderthan-x-hours'],
    Condition['olderthan-x-days'], Condition['olderthan-x-months'],
    Condition['olderthan-x-weeks'], Condition['olderthan-x-years'],
  ].includes(option);

const isDateOptionSetInput = (option: number) =>
  [Condition['in-fiscal-period'], Condition['in-fiscal-year']].includes(option);

export const isDecimalInput =
(columnType: string) => [ 'Decimal', 'FP', 'Currency'].includes(columnType);

export const getInputComponentType = (columnType: string, option: number) => {
  if (columnType.includes('DateAndTime')) {
    if (isFiscalPeriodAndYear(option)) {
      return 'DateFiscalPeriodAndYear';
    }
    else if (isDateTextInput(option)) {
      return 'Text';
    }
    else if (isDateOptionSetInput(option)) {
      return 'DateOptionSet';
    }
    return 'Date';
  }

  if (['OptionSet', 'TwoOptions', 'MultiSelectPicklist', 'Whole.TimeZone'].includes(columnType)) {
    if (isMultipleInput(option, columnType)) return 'Dropdown';

    if (columnType === 'Whole.TimeZone') return 'Number';

    return 'Text';
  }

  if (columnType.includes('Lookup')) {
    if (isMultipleInput(option, columnType)) return 'Lookup';

    return 'Text';
  }

  if (columnType.includes('Whole.Language')) return 'Langauge';

  if (['FP', 'Decimal', 'Whole.None', 'Currency'].includes(columnType)) return 'Number';

  if (['SingleLine.Text', 'Multiple'].includes(columnType)) {
    return 'Text';
  }
};

export const getDateOptionSetOptions = (option: number) =>
  option === Condition['in-fiscal-period'] ? fiscalPeriodOptions
    : Condition['in-fiscal-year'] ? fiscalYearOptions()
      : [];
