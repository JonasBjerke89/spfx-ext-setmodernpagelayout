import * as React from 'react';
import { ISetModernPageLayoutProps } from './ISetModernPageLayoutProps';
import { ISetModernPageLayoutState } from './ISetModernPageLayoutState';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import styles from './SetModernPageLayoutComponent.module.scss';
import { sp } from "@pnp/sp";
import { SPPermission } from "@microsoft/sp-page-context";

export default class SetModernPageLayoutComponent extends React.Component<ISetModernPageLayoutProps, ISetModernPageLayoutState> {
    constructor(props: ISetModernPageLayoutProps) {
        super(props);
        
        this.state = {
            editMode: false,
            showPanel: false
        };
    }

    public componentDidMount(): void {
        // Binding to page mode changes
        const _pushState = () => {
          const _defaultPushState = history.pushState;
          // We need the current this context to update the component its state
          const _self = this;
          return function (data: any, title: string, url?: string | null) {
            // We need to call the in context of the component
            _self.setState({
                editMode: url.indexOf('Mode=Edit') !== -1 && url.indexOf('SitePages') !== -1
            });
    
            // Call the original function with the provided arguments
            // This context is necessary for the context of the history change
            return _defaultPushState.apply(this, [data, title, url]);
          };
        };
        history.pushState = _pushState();
      }

    public render() : React.ReactElement<ISetModernPageLayoutProps> { 
        if((this.state.editMode || window.location.href.indexOf('Mode=Edit') > -1) && this.props.context.pageContext.list.permissions.hasPermission(SPPermission.manageWeb))
        {
            return (
                <div className={styles.btnShow}>
                    <PrimaryButton text="Change modern page layout" iconProps={{ iconName: 'PageSolid' }} onClick={this.setPageLayout.bind(this)} />
                    <Panel
                    isOpen={this.state.showPanel}
                    type={PanelType.smallFixedFar}
                    headerText="Modern Page Layout Options"
                    >
                <Dropdown
                    placeholder="Select a page layout"
                    label="Select a modern page layout:"
                    id="Basicdrop1"
                    className={styles.DropDown}
                    options={[
                        { key: 'Home', text: 'Home Page Layout' },
                        { key: 'Article', text: 'Article Page Layout' },
                        { key: 'SingleWebPartAppPage', text: 'Single WebPart App Page Layout' },
                    ]}
                    onChanged={this.onDropdownChanged.bind(this)}
                    />
                
                    <PrimaryButton text="Save" onClick={this.btnClicked.bind(this)} />  
                    </Panel>
                </div>
            );
        } else
        {
            return (<div />)
        }
    }

    private onDropdownChanged(option: IDropdownOption)
    {
        this.setState(
            {
                pageLayout: option
            }
        );
    }

    private setPageLayout()
    {
        this.setState(
            {
                showPanel: true
            }
        );
    }
    
    private btnClicked(): void {
      sp.web.lists.getById(this.props.context.pageContext.list.id.toString()).items.getById(this.props.context.pageContext.listItem.id).update({PageLayoutType: this.state.pageLayout.key}).then(i => {
        let text = this.state.editMode ? 'We have changed the page layout. We will refresh your page now and keep it in edit mode. Please accept the refresh by clicking "Leave" in a few seconds.' : 'We have changed the page layout. We will refresh your page now.';
        alert(text);
        
        window.location.href = window.location.href;
        });        
    }
}