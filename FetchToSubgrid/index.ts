import * as React from 'react';
import { IInputs, IOutputs } from './generated/ManifestTypes';
import { AppWrapper, IAppWrapperProps } from './components/AppWrapper';
import { DataverseService, IDataverseService } from './services/dataverseService';
import { IColumn } from '@fluentui/react';
import { IColumn } from '@fluentui/react';

export class FetchToSubgrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private _dataverseService: IDataverseService;
  private _primaryEntityName!: string;
  private _fetchXML!: string | null;
  private _columnLayout!: Array<IColumn>;
  private _isDebugMode!: boolean;
  private _baseEnvironmentUrl?: string;
  private _itemsPerPage!: number | null;
  private notifyOutputChanged!: () => void;
  // private _totalNumberOfRecords: number;

  /** General */
  private _context!: ComponentFramework.Context<IInputs>;
  private _notifyOutputChanged!: () => void;
  private _container!: HTMLDivElement;

  public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void): void {
    context.mode.trackContainerResize(true);
    const envVariable = context.parameters.evUrltoCopyRecords.raw;

    this._dataverseService = new DataverseService(context, envVariable);
    this.notifyOutputChanged = notifyOutputChanged;
    this.initVars(context, notifyOutputChanged);
  }

  private initVars(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void): void {
    this._context = context;

    this._notifyOutputChanged = notifyOutputChanged;
    this._isDebugMode = false;
    this._itemsPerPage = 5000;

    // If you want this to break every time you set isDebugMode to true
    //if (this._isDebugMode) {
    //    debugger;  // eslint-disable-line no-debugger
    //}

    debugger;
    // TODO: Validate the input parameters to make sure we get a friendly error instead of weird errors
    var fetchXML: string | null = this._context.parameters.fetchXml.raw;
    var recordIdPlaceholder: string | null = this._context.parameters.RecordIdPlaceholder.raw; // ?? "";

    // This is just the simple control where the subgrid will be placed on the form
    // var controlAnchorField: string | null = this._context.parameters.ControlAnchorField.raw;
    // const recordIdLookupValue: ComponentFramework.EntityReference = this._context.parameters.RecordId.raw[0];
    // Other values if we need them
    let entityId = (<any>this._context.mode).contextInfo.entityId;
    let entityTypeName = (<any>this._context.mode).contextInfo.entityTypeName;
    console.log('🚀 ~ FetchToSubgrid ~ initVars ~ entityTypeName:', entityTypeName);
    let entityDisplayName = (<any>this._context.mode).contextInfo.entityRecordName;
    console.log('🚀 ~ FetchToSubgrid ~ initVars ~ entityDisplayName:', entityDisplayName);
    // This breaks when you use the PCF Test Harness.  Neat!
    try {
      this._baseEnvironmentUrl = (<any>this._context)?.page?.getClientUrl();
    } catch (ex) {
      this._baseEnvironmentUrl = 'https://localhost';
    }
    var recordId: string = entityId; //this._context.parameters.RecordId.raw ?? currentRecordId;
    console.log('🚀 ~ FetchToSubgrid ~ initVars ~ entityId:', entityId);

    // See if we can use an Id from a lookup field specified on the current form
    // Wish we could use the Lookup property type, but doesn't appear to be supported yet
    // https://butenko.pro/2021/03/21/pcf-lookup-attribute-lets-take-look-under-the-hood/
    // So using a hack to get the value from the Xrm.Page.  This is not recommended.
    // TODO: you may need to webapi fetch the id
    // https://github.com/shivuparagi/GenericLookupPCFControl/blob/main/GenericLookupPCFComponent/components/CalloutControlComponent.tsx
    // Look at LoadData function
    var overriddenRecordIdFieldName: string | null = this._context.parameters.OverriddenRecordIdFieldName.raw; // ?? "";
    if (overriddenRecordIdFieldName) {
      try {
        // Hack to get the field value from parent Model Driven App
        // eslint-disable-next-line no-undef
        // @ts-ignore
        // eslint-disable-next-line no-undef
        let tmpLookupField: any = Xrm.Page.getAttribute(overriddenRecordIdFieldName);
        if (tmpLookupField && tmpLookupField.getValue() && tmpLookupField.getValue()[0] && tmpLookupField.getValue()[0].id) {
          recordId = tmpLookupField.getValue()[0].id;
          if (this._isDebugMode) {
            console.log(`overriddenRecordIdFieldName '${overriddenRecordIdFieldName}' value used: ${recordId}.`);
          }
        } else {
          if (this._isDebugMode) {
            console.log(`Could not find id from overriddenRecordIdFieldName '${overriddenRecordIdFieldName}'.`);
          }
        }
        //let control = (<any>this._context)?.page.getControl(overriddenRecordIdFieldName);
        //if (control && control.id){
        //    recordId = control.id;
        //}
      } catch (ex) {
        if (this._isDebugMode) {
          console.log(`Error trying to find id from overriddenRecordIdFieldName '${overriddenRecordIdFieldName}'. ${ex}`);
        }
      }
    }

    // Update FetchXml, replace Record Id Placeholder with an actual Id
    // Grab primary entity Name from FetchXml
    // Test harness always initially passes in "val", so we can skip the following
    if (fetchXML != null && fetchXML != 'val') {
      fetchXML = fetchXML.replace(/"/g, "'");
      this._primaryEntityName = this.getPrimaryEntityNameFromFetchXml(fetchXML);
      // Replace the placeholder
      this._fetchXML = this.replacePlaceholderWithId(fetchXML, recordId, recordIdPlaceholder ?? '');
    }
  }

  private replacePlaceholderWithId(fetchXML: string, recordId: string, recordIdPlaceholder: string): string {
    if (recordId && recordIdPlaceholder) {
      if (fetchXML.indexOf(recordIdPlaceholder) > -1) {
        //return fetchXML.replace(recordIdPlaceholder, recordId); // only replaces first occurrence of string
        return this.replaceAll(fetchXML, recordIdPlaceholder, recordId);
      }
    }
    return fetchXML;
  }

  // Replace ALL occurrences of a string
  private replaceAll(source: string, find: string, replace: string): string {
    // eslint-disable-next-line no-useless-escape
    return source.replace(new RegExp(find.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'), replace);
  }

  private getPrimaryEntityNameFromFetchXml(fetchXml: string): string {
    let primaryEntityName: string = '';
    // @ts-ignore
    let filter = fetchXml.matchAll(/<entity name='(.*?)'>/g).next();
    if (filter && filter.value && filter.value[1]) {
      primaryEntityName = filter.value[1];
    }
    return primaryEntityName;
  }

  public updateView(): React.ReactElement {
    const props: IAppWrapperProps = {
      ...this._dataverseService.getProps(), // Spread existing service props
      fetchXmlOrJson: this._fetchXML, // Pass the fetchXML property
      allocatedWidth: this._context.mode.allocatedWidth || 0, // Add allocatedWidth from context
      default: {
        fetchXml: this._fetchXML, // Default fetchXML value
        pageSize: this._itemsPerPage || 5000, // Default page size
        deleteButtonVisibility: true, // Default visibility for Delete Button
        newButtonVisibility: true, // Default visibility for New Button
      },
    };

    return React.createElement(AppWrapper, props);
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {}
}
