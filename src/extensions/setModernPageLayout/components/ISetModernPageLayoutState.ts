import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export interface ISetModernPageLayoutState
{
    editMode: boolean;
    showPanel: boolean;
    pageLayout?: IDropdownOption;
}