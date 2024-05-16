import * as React from 'react';
import { useState, useCallback, useEffect, useRef } from 'react';
import { Entity, IService } from '../@types/types';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  ISelection,
  ContextualMenu,
  IContextualMenuItem,
  DirectionalHint,
  Stack,
  IComboBoxOption,
} from '@fluentui/react';
import { IDataverseService } from '../services/dataverseService';
import { FilterPopup } from './FilterPopup';
import {
  addFilterIconToColumn,
  getFilterItems,
  removeFilterIconToColumn,
} from '../utilities/filterUtils';
import { checkButtonsAvailable, getColumnName } from '../utilities/fetchXmlUtils';
import { sortColumns } from '../utilities/utils';
import { AttributeType } from '../@types/enums';

interface IListProps extends IService<IDataverseService> {
  entityName: string;
  fetchXml: string | null;
  forceReRender: number;
  updatedFetchXml: string;
  removingAttributeName: React.MutableRefObject<string>;
  currentLinkEntityName: React.MutableRefObject<string>;
  columnType: React.MutableRefObject<string>;
  columns: IColumn[];
  items: Entity[];
  selection: ISelection;
  setSortingData: any;
  setFilteringData: any;
}

export interface IInitialInputValue {
  fieldName: string;
  inputValue: string;
  selectedOption: IComboBoxOption;
}

export const List: React.FC<IListProps> = ({
  _service: dataverseService,
  entityName,
  fetchXml,
  columns,
  forceReRender,
  removingAttributeName,
  currentLinkEntityName,
  columnType,
  items,
  selection,
  setSortingData,
  setFilteringData,
}: IListProps) => {
  const [menuVisible, setMenuVisible] = useState(false);
  const [selectedColumn, setSelectedColumn] = useState<IColumn | null>(null);
  const [target, setTarget] = useState<HTMLElement>();
  const [menuItems, setMenuItems] = useState<IContextualMenuItem[]>([]);
  const [isFilterPopupVisible, setIsFilterPopupVisible] = useState(false);
  const [initialInputValues, setInitialInputValues] = useState<IInitialInputValue[]>([]);

  const columnSortTypeRef = useRef<string>();

  const handleSortItems = (sortType: string) => {
    if (!selectedColumn) return;
    let fieldName = selectedColumn?.className === 'linkEntity' ||
    selectedColumn?.className === 'colIsNotSortable'
      ? selectedColumn?.ariaLabel : selectedColumn?.fieldName;

    if (!fieldName) {
      fieldName = selectedColumn.data?.initialColumnData?._logicalName;
    }

    sortColumns(columns, selectedColumn?.fieldName, sortType === 'descending');
    setSortingData({ fieldName, column: selectedColumn });

    selectedColumn.isSortedDescending = sortType === 'ascending';
    columnSortTypeRef.current = sortType;

    setMenuVisible(false);
  };

  const handleRemoveFilter = () => {
    if (!selectedColumn || !selectedColumn.fieldName) return;
    let { fieldName } = selectedColumn;
    const { attributeType } = selectedColumn.data;
    let columnFieldName = getColumnName(selectedColumn);

    if (selectedColumn.data.initialColumnData) {
      const aliasedFieldName = selectedColumn.data.initialColumnData._logicalName;
      fieldName = aliasedFieldName;
      columnFieldName = aliasedFieldName;
    }

    removingAttributeName.current = selectedColumn.ariaLabel || columnFieldName!;
    currentLinkEntityName.current = selectedColumn.data?.linkEntityName;
    columnType.current = selectedColumn.className!;

    // const matchingValue = initialInputValues.find(val =>
    //   val.fieldName === column.fieldName ||
    //   val.fieldName === column.data?.initialColumnData._logicalName,
    // );

    const matchingValue = initialInputValues.find(val =>
      val.fieldName === columnFieldName,
    );

    const operation = matchingValue?.selectedOption.text;

    if (attributeType === AttributeType.Lookup && operation !== 'Equals') {
      removingAttributeName.current = `${columnFieldName}name`!;
    }

    removeFilterIconToColumn(selectedColumn.ariaLabel || fieldName, columns);

    setFilteringData({
      condition: fieldName,
      column: undefined,
    });

    setSelectedColumn(selectedColumn);
    setMenuVisible(false);
  };

  const onColumnHeaderClick = async (
    ev?: React.MouseEvent<HTMLElement, MouseEvent>,
    column?: IColumn,
  ) => {
    if (!column) return;

    setSelectedColumn(column);
    setTarget(ev?.target as HTMLElement);
    setMenuVisible(true);
  };

  const handleConfirm = (selectedOption: IComboBoxOption, inputValue: string) => {
    if (!selectedColumn || !selectedColumn.fieldName) return;
    const { fieldName } = selectedColumn;
    const { attributeFormat, attributeType } = selectedColumn.data;
    const columnFieldName = getColumnName(selectedColumn);
    let lookUpColumnName = '';

    if (attributeType === AttributeType.Lookup || attributeType === AttributeType.Customer) {
      if (selectedOption.text !== 'Equals') {
        lookUpColumnName = `${columnFieldName}name`;
      }
    }
    else {
      lookUpColumnName = '';
    }

    addFilterIconToColumn(fieldName, columns);
    removingAttributeName.current = '';

    setFilteringData({
      column: selectedColumn,
      fieldName: columnFieldName,
      lookupFieldName: lookUpColumnName,
      attributeFormat,
      selectedOption,
      inputValue,
    });

    const newInputValue = { fieldName, inputValue, selectedOption };

    const matchingValueIndex = initialInputValues.findIndex(val => val.fieldName === fieldName);

    if (matchingValueIndex !== -1) {
      const newValues = [...initialInputValues];
      newValues[matchingValueIndex] = { ...newValues[matchingValueIndex], ...newInputValue };

      setInitialInputValues(newValues);
    }
    else {
      setInitialInputValues([...initialInputValues, newInputValue]);
    }
  };

  const onItemInvoked = useCallback((record?: Entity): void => {
    dataverseService.openRecordForm(entityName, record?.id);
  }, [dataverseService, entityName, fetchXml]);

  useEffect(() => {
    if (!selectedColumn) return;

    const isButtonsAvailable = checkButtonsAvailable(selectedColumn.name, fetchXml!);
    if (!isButtonsAvailable) setMenuVisible(false);

    if (selectedColumn) {
      const items = getFilterItems(
        selectedColumn,
        handleRemoveFilter,
        handleSortItems,
        setIsFilterPopupVisible);

      setMenuItems(items);
    }

  }, [selectedColumn, columnSortTypeRef.current, target, menuVisible]);

  return (
    <Stack>
      <DetailsList
        key={forceReRender}
        columns={columns}
        items={items}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        onItemInvoked={onItemInvoked}
        onColumnHeaderClick={onColumnHeaderClick}
        selection={selection}
      />
      {menuVisible && selectedColumn &&
        <ContextualMenu
          items={menuItems}
          gapSpace={2}
          isBeakVisible={false}
          directionalHint={DirectionalHint.bottomLeftEdge}
          onDismiss={() => setMenuVisible(false)}
          target={target}
        />
      }
      {isFilterPopupVisible && target && selectedColumn &&
      <FilterPopup
        column={selectedColumn}
        target={target}
        hideFilterMenu={() => setIsFilterPopupVisible(false)}
        handleConfirm={handleConfirm}
        dataverseService={dataverseService}
        initialInputValues={initialInputValues}
        entityName={entityName}
      />
      }
    </Stack>
  );
};
