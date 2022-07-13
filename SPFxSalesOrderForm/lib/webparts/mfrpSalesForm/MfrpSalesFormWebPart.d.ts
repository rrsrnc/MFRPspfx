import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IMfrpSalesFormWebPartProps {
    description: string;
}
export default class MfrpSalesFormWebPart extends BaseClientSideWebPart<IMfrpSalesFormWebPartProps> {
    render(): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=MfrpSalesFormWebPart.d.ts.map