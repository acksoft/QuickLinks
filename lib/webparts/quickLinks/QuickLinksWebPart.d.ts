import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
import { IQuickLinksWebPartProps } from './IQuickLinksWebPartProps';
import 'jquery';
export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
    render(): void;
    protected readonly dataVersion: Version;
    private getLinkData();
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
