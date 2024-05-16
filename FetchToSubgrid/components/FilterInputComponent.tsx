import * as React from 'react';
import { Dispatch, FC, FormEvent, SetStateAction, useEffect } from 'react';
import {
  TextField,
  ComboBox,
  TagPicker,
  DatePicker,
  IComboBox,
  IComboBoxOption,
  IColumn,
  ITag,
} from '@fluentui/react';
import { AttributeType, WholeNumberType } from '../@types/enums';
import { getDateFormatWithHyphen } from '../utilities/dateTimeUtils';
import { IDataverseService } from '../services/dataverseService';
import { IInitialInputValue } from './List';

interface FilterInputComponentProps {
  dataverseService: IDataverseService;
  column: IColumn;
  inputValue: string;
  entityName: string;
  inputErrorMessage: string;
  initialInputValues: IInitialInputValue[];
  selectedOption: IComboBoxOption;
  setInputValue: Dispatch<SetStateAction<string>>;
  setInputErrorMessage: Dispatch<SetStateAction<string>>;
  onTextChange: (
    event: FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
    ) => void;
}

export const FilterInputComponent: FC<FilterInputComponentProps> = ({
  dataverseService,
  column,
  inputValue,
  entityName,
  inputErrorMessage,
  selectedOption,
  initialInputValues,
  setInputErrorMessage,
  setInputValue,
  onTextChange,
}: FilterInputComponentProps) => {
  const [isValidNumber, setIsValidNumber] = React.useState<boolean>(true);
  const [lookups, setLookups] = React.useState<ITag[]>([]);
  const [options, setOptions] = React.useState<IComboBoxOption[]>([]);

  function findNameByKey(key: string): string | null {
    for (let i = 0; i < lookups.length; i++) {
      const item = lookups[i];

      if (item.key === key) {
        return item.name;
      }
    }
    return null;
  }

  useEffect(() => {
    const getInitialValue = async () => {
      if (!initialInputValues) return;

      const matchingInputValue: IInitialInputValue | undefined = initialInputValues.find(
        (inputValue: any) => inputValue.fieldName === column.fieldName,
      );

      const linkEntityName = column?.data?.linkEntityName;

      if (column.data.attributeType === AttributeType.TwoOptions) {
        const result = await dataverseService.getDropdownOptions(
          column.ariaLabel! || column.fieldName!,
          'BooleanAttributeMetadata',
          linkEntityName || entityName,
          true,
        );
        setOptions(result.options);
      }

      else if (column.data.attributeType === AttributeType.MultiselectPickList) {
        const fieldName = column.fieldName!.startsWith('alias')
          ? column.data?.initialColumnData?._logicalName : column.fieldName!;

        const linkEntityName = column?.data?.linkEntityName;

        const result = await dataverseService.getDropdownOptions(
          column.ariaLabel! || fieldName,
          'MultiSelectPicklistAttributeMetadata',
          linkEntityName || entityName,
          false,
        );

        setOptions(result.options);
      }

      else if (column.data.attributeType === AttributeType.PickList) {
        const fieldName = column.fieldName!.startsWith('alias')
          ? column.data?.initialColumnData?._logicalName : column.fieldName!;

        const linkEntityName = column?.data?.linkEntityName;
        const result = await dataverseService.getDropdownOptions(
          column.ariaLabel || fieldName,
          'PicklistAttributeMetadata',
          linkEntityName || entityName,
          false,
        );

        setOptions(result.options);
      }

      else if (column.data.attributeType === AttributeType.Status) {
        const fieldName = column.fieldName!.startsWith('alias')
          ? column.data?.initialColumnData?._logicalName : column.fieldName!;

        const linkEntityName = column?.data?.linkEntityName;
        const result = await dataverseService.getDropdownOptions(
          column.ariaLabel! || fieldName,
          'StatusAttributeMetadata',
          linkEntityName || entityName,
          false,
        );

        setOptions(result.options);
      }
      else if (column.data.attributeType === AttributeType.StateCode) {
        const fieldName = column.fieldName!.startsWith('alias')
          ? column.data?.initialColumnData?._logicalName : column.fieldName!;

        const linkEntityName = column?.data?.linkEntityName;
        const result = await dataverseService.getDropdownOptions(
          column.ariaLabel! || fieldName,
          'StateAttributeMetadata',
          linkEntityName || entityName,
          false,
        );

        setOptions(result.options);
      }

      if (matchingInputValue && column.data.attributeType === AttributeType.Lookup &&
            matchingInputValue.selectedOption.text === 'Equals'
      ) {
        const item = await dataverseService.getRecordById(
          column.data.fieldEntityName,
          matchingInputValue.inputValue,
        );

        setLookups(item);
      }
      if (matchingInputValue && column.data.attributeType === AttributeType.Customer &&
           matchingInputValue.selectedOption.text === 'Equals'
      ) {
        const option = await dataverseService.getContactOrAccountById(
          matchingInputValue.inputValue);

        setLookups(option);
      }
      else if (matchingInputValue && column.data.attributeFormat === WholeNumberType.Language &&
        matchingInputValue.selectedOption.text === 'Equals'
      ) {
        const option = await dataverseService.getProvisionedLanguages();

        setLookups(option);
      }

      if (
        matchingInputValue &&
        column.isFiltered &&
        column.fieldName === matchingInputValue.fieldName
      ) {

        setInputValue(matchingInputValue.inputValue);
      }
      else {
        setInputValue('');
      }
    };
    getInitialValue();
  }, []);

  const lookupSelectedOptions: ITag[] = inputValue
    ? inputValue.split(',').map(val => ({
      name: findNameByKey(val) || '',
      key: val,
    }))
    : [];

  const onDateChange = (date: Date | null | undefined) => {
    if (date === null || date === undefined) {
      setInputValue('0');
      setInputErrorMessage('Date is invalid. Example: 2/17/1999.');
    }
    else {
      setInputValue(getDateFormatWithHyphen(date));
      setInputErrorMessage('');
    }
  };

  const handleTwoOptionChange = (
    event: FormEvent<IComboBox>,
    option?: IComboBoxOption,
  ): void => {
    if (option && option.key) {
      setInputValue(option.key as string);
    }
  };

  const onLookupChange = (items?: ITag[] | undefined) => {
    if (items && items.length > 0 && items[0].key) {
      setInputValue(items[0].key as string);
      setInputErrorMessage('');
    }
    else {
      setInputValue('');
    }
  };

  const onNumberChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string,
  ) => {
    if (!isNaN(Number(newValue))) {
      setInputValue(newValue || '');
      setInputErrorMessage('');
      setIsValidNumber(true);
    }
    else {
      setInputValue('');
      setInputErrorMessage('Value must be a valid whole number');
      setIsValidNumber(false);
    }
  };

  const listContainsTagList = (tag: ITag, tagList?: ITag[]) => {
    if (!tagList || !tagList.length || tagList.length === 0) {
      return false;
    }
    return tagList.some(compareTag => compareTag.key === tag.key);
  };

  const filterSuggestedTags = (filterText: string, tagList: ITag[] | undefined): ITag[] => {
    if (filterText.length === 0) return [];

    return lookups.filter((tag: ITag) => {
      if (tag.name === null) return false;

      return tag.name.toLowerCase().includes(filterText.toLowerCase()) &&
      !listContainsTagList(tag, tagList);
    });
  };

  const initialValues = async (selectedItems: ITag[] | undefined): Promise<ITag[]> => {
    let result: ITag[] = [];

    if (selectedItems?.length === 1) return [];
    if (column.data.attributeType === AttributeType.Number &&
      column.data.attributeFormat === WholeNumberType.Language) {
      result = await dataverseService.getProvisionedLanguages();
      setLookups(result);
    }

    if (column.data.attributeType === AttributeType.Lookup) {
      result = await dataverseService.getLookupOptions(column.data.fieldEntityName);
      setLookups(result);
    }

    if (column.data.attributeType === AttributeType.Customer) {
      result = await dataverseService.getAccountsAndContactsOptions();
      setLookups(result);
    }

    const initialItems =
      selectedItems === undefined || selectedItems.length < 1
        ? result
        : result.filter((option: ITag) => !selectedItems.includes(option));

    if (initialItems.length > 300) {
      return initialItems.slice(0, 300);
    }
    return initialItems;
  };

  switch (column.data.attributeType) {
    case AttributeType.DateTime:
      return (
        <DatePicker
          className="filterInput"
          value={inputValue ? new Date(inputValue) : undefined}
          onSelectDate={onDateChange}
          formatDate={(date?: Date) => (date ? date.toLocaleDateString() : '')}
        />
      );
    case AttributeType.TwoOptions:
      return (
        <ComboBox
          className="filterInput"
          options={options}
          onChange={handleTwoOptionChange}
          selectedKey={inputValue}
        />
      );
    case AttributeType.Status:
      return (
        <ComboBox
          className="filterInput"
          options={options}
          onChange={handleTwoOptionChange}
          selectedKey={inputValue}
        />
      );
    case AttributeType.StateCode:
      return (
        <ComboBox
          className="filterInput"
          options={options}
          onChange={handleTwoOptionChange}
          selectedKey={inputValue}
        />
      );

    case AttributeType.MultiselectPickList:
      return (
        <ComboBox
          className="filterInput"
          options={options}
          onChange={handleTwoOptionChange}
          selectedKey={inputValue}
        />
      );
    case AttributeType.PickList:
      return (
        <ComboBox
          className="filterInput"
          options={options}
          onChange={handleTwoOptionChange}
          selectedKey={inputValue}
        />
      );
    case AttributeType.Lookup:
    case AttributeType.Customer:
      if (selectedOption.text === 'Equals') {
        return (
          <TagPicker
            selectedItems={lookupSelectedOptions}
            className="filterInput"
            onChange={onLookupChange}
            onResolveSuggestions={filterSuggestedTags}
            resolveDelay={1000}
            onEmptyResolveSuggestions={initialValues}
          />
        );
      }
      return (
        <TextField
          className="filterInput"
          value={inputValue}
          onChange={onTextChange}
          errorMessage={inputErrorMessage}
        />
      );

    case AttributeType.Number:
      if (column.data.attributeFormat === '3') {
        return (
          <TagPicker
            selectedItems={lookupSelectedOptions}
            className="filterInput"
            onChange={onLookupChange}
            onResolveSuggestions={filterSuggestedTags}
            resolveDelay={1000}
            onEmptyResolveSuggestions={initialValues}
          />
        );
      }
      return (
        <TextField
          className="filterInput"
          value={inputValue}
          onChange={onNumberChange}
          errorMessage={!isValidNumber ? 'Please enter a valid number' : inputErrorMessage}
        />
      );

    case AttributeType.DecimalNumber:
    case AttributeType.Money:
    case AttributeType.FloatNumber:
      return (
        <TextField
          className="filterInput"
          value={inputValue}
          onChange={onNumberChange}
          errorMessage={!isValidNumber ? 'Please enter a valid number' : inputErrorMessage}
        />
      );
    default:
      return (
        <TextField
          className="filterInput"
          value={inputValue}
          onChange={onTextChange}
          errorMessage={inputErrorMessage}
        />
      );
  }
};
