import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";

export interface ISetModernPageLayoutProps
{
    context: ApplicationCustomizerContext;
    listId: string;
    itemId: number;
}