/* global HTMLCollectionOf, NodeListOf */
import { IColumn } from '@fluentui/react';
import { Dictionary, EntityAttribute, OrderInFetchXml } from '../@types/types';
import { formatCondition } from './filterUtils';

export const changeAliasNames = (fetchXml: string) => {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(fetchXml, 'text/xml');

  const attributeElements = xmlDoc.querySelectorAll('attribute[alias]');

  attributeElements.forEach((attributeElement, index) => {
    const newAliasValue = `alias${index}`;
    attributeElement?.setAttribute('alias', newAliasValue);
  });

  return xmlDoc.documentElement.outerHTML;
};

export const addPagingToFetchXml = (
  fetchXml: string,
  pageSize: number,
  currentPage: number): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');
  const top: string | null | undefined = xmlDoc.querySelector('fetch')?.getAttribute('top');

  const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];
  fetch?.removeAttribute('count');
  fetch?.removeAttribute('page');
  fetch?.removeAttribute('top');

  let recordsPerPage = pageSize;

  if (top && currentPage * recordsPerPage > Number(top)) {
    recordsPerPage = Number(top) % pageSize;
  }

  fetch?.setAttribute('page', `${currentPage}`);
  fetch?.setAttribute('count', `${recordsPerPage}`);

  const newFetchChangedAliases = changeAliasNames(new XMLSerializer().serializeToString(xmlDoc));

  return newFetchChangedAliases;
};

export const getEntityNameFromFetchXml = (fetchXml: string): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  return xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
};

export const getOrderInFetchXml = (fetchXml: string): OrderInFetchXml | null => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: XMLDocument = parser.parseFromString(fetchXml, 'text/xml');
  const entityOrder: any = xmlDoc.querySelector('entity > order');
  const linkEntityOrder: any = xmlDoc.querySelectorAll('link-entity > order');

  if (entityOrder && linkEntityOrder.length > 0 || linkEntityOrder.length > 1) return null;

  const order = entityOrder || linkEntityOrder[0];
  if (!order) return null;

  const isLinkEntity = linkEntityOrder[0] !== null;

  const descending: boolean = order.attributes.descending?.value === 'true';
  const attribute: string = order.attributes.attribute?.value;

  return {
    [attribute]: descending,
    isLinkEntity,
  };
};

// const linkEntities = xmlDoc.getElementsByTagName('link-entity');

// let linkEntity;

// let linkEntityFilter;

// for (let i = 0; i < linkEntities.length; i++) {
//   const currentLinkEntity = linkEntities[i];

//   currentLinkEntity?.removeAttribute('type');
//   const name = currentLinkEntity.getAttribute('name');

//   if (name === filteringData.column?.data?.linkEntityName) {
//     linkEntity = currentLinkEntity;
//     const [firstFilter] = Array.from(currentLinkEntity.getElementsByTagName('filter'));

//     linkEntityFilter = firstFilter;
//     break;
//   }
// }

export const isAggregate = (fetchXml: string | null): boolean => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const aggregate = xmlDoc.getElementsByTagName('fetch')?.[0]?.getAttribute('aggregate') ?? '';
  if (aggregate === 'true') return true;

  return false;
};

export const addOrderToFetch = (
  fetchXml: string | null,
  sortingData: { fieldName: string, column?: IColumn }): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const entity: Element = xmlDoc.getElementsByTagName('entity')[0];
  // const linkEntity: Element = xmlDoc.getElementsByTagName('link-entity')[0];
  const linkEntities = xmlDoc.getElementsByTagName('link-entity');
  let linkEntity;

  for (let i = 0; i < linkEntities.length; i++) {
    const currentLinkEntity = linkEntities[i];

    const name = currentLinkEntity.getAttribute('name');

    if (name === sortingData.column?.data?.linkEntityName) {
      linkEntity = currentLinkEntity;
      break;
    }
  }

  const entityOrder: Element = entity?.getElementsByTagName('order')[0];
  const linkOrder = linkEntity?.getElementsByTagName('order')[0];

  if (linkOrder && linkOrder.parentNode === linkEntity) {
    linkEntity.removeChild(linkOrder);
  }
  else if (entityOrder && entityOrder.parentNode === entity) {
    entity.removeChild(entityOrder);
  }

  const hasAggregate = isAggregate(fetchXml ?? '');

  const newOrder: HTMLElement = xmlDoc.createElement('order');

  if (sortingData.column && sortingData.column.fieldName &&
    hasAggregate && sortingData.column.fieldName.startsWith('alias')) {

    newOrder?.setAttribute('alias', `${sortingData?.column?.fieldName}`);
    newOrder?.setAttribute('descending', `${!sortingData.column?.isSortedDescending}`);
  }
  else {
    newOrder?.setAttribute('attribute', `${sortingData.fieldName}`);
    newOrder?.setAttribute('descending', `${!sortingData.column?.isSortedDescending}`);
  }

  if (sortingData.column?.className === 'entity') {
    entity.appendChild(newOrder);
  }
  else if (sortingData.column?.className === 'colIsNotSortable') {
    if (sortingData.column.data.linkEntityName) {
      linkEntity?.appendChild(newOrder);
    }
    else {
      entity.appendChild(newOrder);
    }
  }
  else {
    linkEntity?.appendChild(newOrder);
  }

  return new XMLSerializer().serializeToString(xmlDoc);
};

export const addFilterToFetch = (
  fetchXml: string | null,
  filteringData: any,
): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const entity: Element = xmlDoc.getElementsByTagName('entity')[0];
  const linkEntities = xmlDoc.getElementsByTagName('link-entity');

  let linkEntity;

  let linkEntityFilter;

  for (let i = 0; i < linkEntities.length; i++) {
    const currentLinkEntity = linkEntities[i];

    currentLinkEntity?.removeAttribute('type');
    const name = currentLinkEntity.getAttribute('name');

    if (name === filteringData.column?.data?.linkEntityName) {
      linkEntity = currentLinkEntity;
      const [firstFilter] = Array.from(currentLinkEntity.getElementsByTagName('filter'));

      linkEntityFilter = firstFilter;
      break;
    }
  }

  const entityFilters = entity?.getElementsByTagName('filter');
  let entityFilter;

  if (entityFilters) {
    for (let i = 0; i < entityFilters.length; i++) {
      const filter = entityFilters[i];
      if (filter.parentNode === entity) {
        entityFilter = filter;
        break;
      }
    }
  }

  let existingFilter: Element | null = null;

  let operatorName = '';
  let conditionValue = '';

  if (typeof filteringData?.selectedOption?.key === 'number') {
    const condition = formatCondition(
      filteringData?.selectedOption?.key,
      filteringData.inputValue,
      filteringData.attributeFormat,
    );
    operatorName = condition.conditionOperator;
    conditionValue = condition.value;
  }

  if (entityFilter) {
    const conditions: HTMLCollectionOf<Element> = entityFilter.getElementsByTagName('condition');

    for (let i = 0; i < conditions.length; i++) {
      const condition: Element = conditions[i];

      if (filteringData.lookupFieldName) {
        if (condition.getAttribute('attribute') === filteringData.fieldName ||
            condition.getAttribute('attribute') === filteringData.lookupFieldName
        ) {
          condition.setAttribute('attribute', filteringData.lookupFieldName);
          condition.setAttribute('value', conditionValue);
          condition.setAttribute('operator', operatorName);
          existingFilter = entityFilter;
          break;
        }
      }
      else if (condition.getAttribute('attribute') === filteringData.fieldName ||
          condition.getAttribute('attribute') === `${filteringData.fieldName}name`
      ) {
        condition.setAttribute('attribute', filteringData.fieldName);
        condition.setAttribute('value', conditionValue);
        condition.setAttribute('operator', operatorName);

        existingFilter = entityFilter;
        break;
      }
      else if (filteringData.column) {
        if (filteringData.column.data?.initialColumnData?._logicalName ===
          condition.getAttribute('attribute')) {

          condition.setAttribute('attribute',
            filteringData.column.data?.initialColumnData?._logicalName);

          condition.setAttribute('value', conditionValue);
          condition.setAttribute('operator', operatorName);
          existingFilter = entityFilter;
          break;
        }
        else if (filteringData.column.className !== 'linkEntity') {
          const attributeName = filteringData.column.data?.initialColumnData?._logicalName ||
          filteringData.fieldName;

          condition.setAttribute('attribute', attributeName);
          condition.setAttribute('value', conditionValue);
          condition.setAttribute('operator', operatorName);
          existingFilter = entityFilter;
          break;
        }
      }
    }
  }

  if (!existingFilter) { // create new filter attribute
    const newFilter: HTMLElement = xmlDoc.createElement('filter');
    newFilter.setAttribute('type', 'and');

    const newCondition: HTMLElement = xmlDoc.createElement('condition');
    newCondition.setAttribute('operator', operatorName);
    newCondition.setAttribute('value', conditionValue);

    newCondition.setAttribute('attribute', filteringData.lookupFieldName
      ? filteringData.lookupFieldName
      : filteringData.fieldName);

    if (filteringData.column?.className === 'entity') {
      newFilter.appendChild(newCondition);
      entity.appendChild(newFilter);
    }
    else if (filteringData.column?.className === 'linkEntity' && linkEntity) {
      if (linkEntityFilter) {
        linkEntityFilter.appendChild(newCondition);
      }
      else {
        newFilter.appendChild(newCondition);
        linkEntity.appendChild(newFilter);
      }
    }
    else if (filteringData.column?.className === 'colIsNotSortable') {
      if (filteringData.column.ariaLabel) {
        newCondition.setAttribute('attribute', filteringData.column.ariaLabel);
      }
      else {
        newCondition.setAttribute('attribute',
          filteringData.column.data?.initialColumnData?._logicalName);
      }

      const linkEntityName = filteringData.column.data?.linkEntityName;

      if (linkEntityName) {
        newFilter.appendChild(newCondition);
        linkEntity!.appendChild(newFilter);
      }
      else {
        newFilter.appendChild(newCondition);
        entity.appendChild(newFilter);
      }

    }
  }

  return new XMLSerializer().serializeToString(xmlDoc);
};

const removeFilterConditions = (parentElement: Element, attributeName: string) => {
  const filters: HTMLCollectionOf<Element> = parentElement.getElementsByTagName('filter');

  for (let i = 0; i < filters.length; i++) {
    const filter: Element = filters[i];
    const conditions: HTMLCollectionOf<Element> = filter.getElementsByTagName('condition');

    for (let j = 0; j < conditions.length; j++) {
      const condition: Element = conditions[j];

      if (condition.getAttribute('attribute') === attributeName) {
        filter.removeChild(condition);
      }
    }
  }
};

export const removeColumnFilter = (
  fetchXml: string,
  attributeName: string,
  currentLinkEntityName: string,
  columnType: string): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const entity: Element | null = xmlDoc.getElementsByTagName('entity')[0];
  // const linkEntity: Element | null = xmlDoc.getElementsByTagName('link-entity')[0];
  const linkEntities = xmlDoc.getElementsByTagName('link-entity');
  let linkEntity;

  for (let i = 0; i < linkEntities.length; i++) {
    const currentLinkEntity = linkEntities[i];

    const name = currentLinkEntity.getAttribute('name');

    if (name === currentLinkEntityName) {
      linkEntity = currentLinkEntity;
    }
  }

  if (entity && columnType === 'entity') {
    removeFilterConditions(entity, attributeName);
  }

  if (linkEntity && columnType === 'linkEntity') {
    removeFilterConditions(linkEntity, attributeName);
  }
  else if (columnType === 'colIsNotSortable') {
    removeFilterConditions(entity, attributeName);
  }

  const modifiedFetchXml = new XMLSerializer().serializeToString(xmlDoc);

  return modifiedFetchXml;
};

export const getColumnName = (column: IColumn): string => {
  const { fieldName, className, ariaLabel } = column;
  const isLinkEntity = className === 'linkEntity';
  const columnFieldName = isLinkEntity ? ariaLabel : fieldName;

  return columnFieldName || '';
};

export const getLinkEntitiesNamesFromFetchXml = (
  fetchXml: string): Dictionary<EntityAttribute[]> => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const linkEntities: HTMLCollectionOf<Element> = xmlDoc.getElementsByTagName('link-entity');
  const linkEntityData: Dictionary<EntityAttribute[]> = {};

  Array.prototype.slice.call(linkEntities).forEach(linkentity => {
    const entityName: string = linkentity.attributes['name'].value;
    const entityAttributes: EntityAttribute[] = [];

    const attributesSelector = `link-entity[name="${entityName}"] > attribute`;
    const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributesSelector);
    const linkEntityAlias: string | undefined = linkentity.attributes['alias']?.value;

    Array.prototype.slice.call(attributes).map(attr => {
      const attributeAlias: string = attr.attributes['alias']?.value ?? '';

      entityAttributes.push({
        linkEntityAlias,
        name: attr.attributes.name.value,
        attributeAlias,
      });
    });

    linkEntityData[entityName] = entityAttributes;
  });

  return linkEntityData;
};

export const getFetchXmlAttributesData = (
  fetchXml: string | null,
  isAggregate: boolean): string[] => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const entityName = xmlDoc.getElementsByTagName('entity')?.[0]?.getAttribute('name') ?? '';
  const attributesFieldNames: string[] = [];

  const attributeSelector = `entity[name="${entityName}"] > attribute`;
  const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributeSelector);

  Array.prototype.slice.call(attributes).map(attr => {
    if (isAggregate) {
      attributesFieldNames.push(attr.attributes.name.value);
    }
    else {
      attributesFieldNames.push(attr.attributes.alias?.value);
    }
  });

  return attributesFieldNames;
};

export const checkButtonsAvailable = (columnName: string, fetchXml: string): boolean => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const attributeWithAlias = xmlDoc.querySelector(`attribute[alias="${columnName}"]`);

  if (attributeWithAlias) {
    const aggregateElement = attributeWithAlias.getAttribute('aggregate');
    if (aggregateElement) {
      return false;
    }
  }

  return true;
};

export const addAttributeIdInFetch = (fetchXml: string, entityName: string): string => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const fetchElement: Element | null = xmlDoc.querySelector('fetch');
  const distinctAttribute: string | null | undefined = fetchElement?.getAttribute('distinct');

  if (distinctAttribute !== 'true') return fetchXml;

  const entityElement: Element | null = xmlDoc.querySelector(`entity[name="${entityName}"]`);

  if (entityElement) {
    const entityIdAttribute = `${entityName}id`;

    const existingAttributeId: Element | null = entityElement.querySelector(
      `attribute[name="${entityIdAttribute}"]`);

    const aggregate = xmlDoc.getElementsByTagName('fetch')?.[0]?.getAttribute('aggregate') ?? '';

    if (!existingAttributeId) {
      const newAttributeElement: Element = xmlDoc.createElement('attribute');
      newAttributeElement.setAttribute('name', entityIdAttribute);
      if (aggregate === 'true') {
        newAttributeElement.setAttribute('groupby', 'true');
        newAttributeElement.setAttribute('alias', 'Id');
      }

      entityElement.appendChild(newAttributeElement);
    }
  }
  const newFetchXml: string = new XMLSerializer().serializeToString(xmlDoc);
  return newFetchXml;
};

export const getLinkEntityAggregateAliasNames = (fetchXml: string, i: number): string[] => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml, 'text/xml');

  const aggregateAttrNames: string[] = [];

  const entityName = xmlDoc.getElementsByTagName('link-entity')?.[i]?.getAttribute('name') ?? '';
  const attributeSelector = `link-entity[name="${entityName}"] > attribute`;
  const attributes: NodeListOf<Element> = xmlDoc.querySelectorAll(attributeSelector);

  Array.prototype.slice.call(attributes).map(attr => {
    aggregateAttrNames.push(attr.attributes.alias?.value);
  });

  return aggregateAttrNames;
};

export const getTopInFetchXml = (fetchXml: string | null): number => {
  if (!fetchXml) return 0;
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');
  const fetch: Element = xmlDoc.getElementsByTagName('fetch')?.[0];

  const top: string | null = fetch?.getAttribute('top');

  return Number(top) || 0;
};

export const getFetchXmlParserError = (fetchXml: string | null): string | null => {
  const parser: DOMParser = new DOMParser();
  const xmlDoc: Document = parser.parseFromString(fetchXml ?? '', 'text/xml');

  const parserError: Element | null = xmlDoc.querySelector('parsererror');
  if (!parserError) return null;

  const errorMessage: string | null = parserError.querySelector('div')?.innerText ?? null;
  return errorMessage;
};
