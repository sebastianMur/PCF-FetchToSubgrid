import * as React from 'react';
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { IColumn, Stack } from '@fluentui/react';
import { Entity, IItemsData } from '../@types/types';
import { dataSetStyles } from '../styles/comandBarStyles';
import { CommandBar } from './ComandBar';
import { List } from './List';
import { Footer } from './Footer';
import { useSelection } from '../hooks/useSelection';
import { IDataverseService } from '../services/dataverseService';
import { hashCode, getLinkableItems, getSortedColumns } from '../utilities/utils';
import {
  addFilterToFetch,
  addOrderToFetch,
  getEntityNameFromFetchXml,
  removeColumnFilter } from '../utilities/fetchXmlUtils';

export interface IFetchToSubgridProps {
  fetchXml: string | null;
  _service: IDataverseService;
  pageSize: number;
  deleteButtonVisibility: boolean;
  newButtonVisibility: boolean;
  allocatedWidth: number;
  error?: Error;
  setIsLoading: (isLoading: boolean) => void;
  setError: (error?: Error | undefined) => void;
}

export const FetchToSubgrid: React.FC<IFetchToSubgridProps> = props => {
  const {
    _service: dataverseService,
    deleteButtonVisibility,
    newButtonVisibility,
    pageSize,
    allocatedWidth,
    fetchXml,
    setIsLoading,
    setError,
  } = props;

  const [items, setItems] = useState<Entity[]>([]);
  const [currentPage, setCurrentPage] = useState(1);
  const [isDialogAccepted, setDialogAccepted] = useState(false);
  const [columns, setColumns] = React.useState<IColumn[]>([]);

  const listInputsHashCode = useRef(-1);
  const totalRecordsCount = useRef(0);
  const newFilteredFetchXml = useRef('');
  const updatedFetchXml = useRef('');
  const removingAttributeName = useRef('');
  const columnType = useRef('');

  const [sortingData, setSortingData] = useState({ fieldName: '', column: undefined });
  const [filteringData, setFilteringData] = useState({ lookupFieldName: '', column: undefined });

  const entityName = getEntityNameFromFetchXml(fetchXml ?? '');

  const { selection, selectedRecordIds } = useSelection();
  let isButtonActive = selectedRecordIds.length > 0;

  const determineFetchXml = (
    fetchXmlToUse: string | null,
    newFetchXml: string) => sortingData.column === undefined &&
    filteringData.column === undefined
    ? fetchXmlToUse : newFetchXml;

  const handleFiltering = async () => {
    if (newFilteredFetchXml.current) {
      newFilteredFetchXml.current = addFilterToFetch(
        newFilteredFetchXml.current || fetchXml, filteringData);
    }
    else {
      newFilteredFetchXml.current = addFilterToFetch(fetchXml, filteringData);
    }

    if (filteringData.column) {
      totalRecordsCount.current = await dataverseService.getRecordsCount(
        newFilteredFetchXml.current);
    }
    else {
      totalRecordsCount.current = await dataverseService.getRecordsCount(fetchXml ?? '');
    }

    return newFilteredFetchXml.current;
  };

  useLayoutEffect(() => {
    if (allocatedWidth === -1) return;
    listInputsHashCode.current = hashCode(`${allocatedWidth}${fetchXml}`);
  }, [allocatedWidth, columns]);

  useEffect(() => setCurrentPage(1), [pageSize, fetchXml]);

  useEffect(() => {
    const fetchColumns = async () => {
      try {
        if (allocatedWidth === -1) return;

        const filteredColumns = await getSortedColumns(fetchXml, allocatedWidth, dataverseService);
        setColumns(filteredColumns);
      }
      catch (error: any) {
        setError(error);
      }
    };

    fetchColumns();
  }, [fetchXml, allocatedWidth]);

  useEffect(() => {
    const fetchItems = async () => {
      isButtonActive = false;
      setIsLoading(true);

      if (isDialogAccepted) return;
      try {
        totalRecordsCount.current = await dataverseService.getRecordsCount(fetchXml ?? '');

        let newFetchXml = await handleFiltering();

        if (sortingData.column === undefined && filteringData.column !== undefined) {
          newFetchXml = newFilteredFetchXml.current;
        }

        if (removingAttributeName.current) {
          newFilteredFetchXml.current = removeColumnFilter(
            newFilteredFetchXml.current,
            removingAttributeName.current,
            columnType.current,
          );
        }

        if (sortingData.column !== undefined) {
          newFetchXml = addOrderToFetch(newFilteredFetchXml.current, sortingData);
        }

        const data: IItemsData = {
          fetchXml: determineFetchXml(newFilteredFetchXml.current, newFetchXml),
          pageSize,
          currentPage,
        };

        const linkableItems: Entity[] = await getLinkableItems(data, dataverseService);

        setItems(linkableItems);
      }
      catch (error: any) {
        setError(error);
      }
      setIsLoading(false);
    };

    fetchItems();
  }, [sortingData, filteringData]);

  useEffect(() => {
    const fetchItems = async () => {
      isButtonActive = false;
      setIsLoading(true);

      try {
        totalRecordsCount.current = await dataverseService.getRecordsCount(fetchXml ?? '');

        const data: IItemsData = {
          fetchXml: sortingData.column === undefined ? fetchXml : newFilteredFetchXml.current,
          pageSize,
          currentPage,
        };

        const linkableItems: Entity[] = await getLinkableItems(data, dataverseService);

        setItems(linkableItems);
      }
      catch (error: any) {
        setError(error);
      }
      setIsLoading(false);
    };

    fetchItems();
  }, [currentPage, isDialogAccepted, fetchXml, pageSize]);

  return <>
    <Stack horizontal horizontalAlign="end" className={dataSetStyles.buttons}>
      <CommandBar
        _service={dataverseService}
        entityName={entityName}
        selectedRecordIds={selectedRecordIds}
        setDialogAccepted={setDialogAccepted}
        isButtonActive={isButtonActive}
        deleteButtonVisibility={deleteButtonVisibility}
        newButtonVisibility={newButtonVisibility}
      />
    </Stack>

    <List
      _service={dataverseService}
      entityName={entityName}
      forceReRender={listInputsHashCode.current}
      updatedFetchXml={updatedFetchXml.current}
      removingAttributeName={removingAttributeName}
      columnType={columnType}
      fetchXml={fetchXml}
      selection={selection}
      columns={columns}
      items={items}
      setSortingData={setSortingData}
      setFilteringData={setFilteringData}
    />

    <Footer
      pageSize={pageSize}
      selectedItemsCount={selectedRecordIds.length}
      totalRecordsCount={totalRecordsCount.current}
      currentPage={currentPage}
      setCurrentPage={setCurrentPage}
    />
  </>;
};
