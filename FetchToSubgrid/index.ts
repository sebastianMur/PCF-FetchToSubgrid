import { IInputs, IOutputs } from './generated/ManifestTypes';
import { FetchSubgrid, IFetchSubgridProps } from './Components/FetchSubgrid';
import * as React from 'react';
import FetchService from './Services/CrmService';

export class FetchToSubgrid implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;

    constructor() { }

    public init(
      context: ComponentFramework.Context<IInputs>,
      notifyOutputChanged: () => void,
    ): void {
      this.notifyOutputChanged = notifyOutputChanged;
      FetchService.setContext(context);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
      const props: IFetchSubgridProps = {
        defaultFetchXml: context.parameters.defaultFetchXmlProperty.raw };

      return React.createElement(
        FetchSubgrid, props,
      );
    }

    public getOutputs(): IOutputs {
      return {};
    }

    public destroy(): void {
    }
}
