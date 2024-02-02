import * as React from 'react';
// import styles from './DynamicDropdown.module.scss';
import { IDynamicDropdownProps } from './IDynamicDropdownProps';
import { Dropdown, IDropdownOption, PrimaryButton } from 'office-ui-fabric-react';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { Guid } from '@microsoft/sp-core-library';
import { Dialog } from '@microsoft/sp-dialog';
// import { escape } from '@microsoft/sp-lodash-subset';
import {sp,Web} from '@pnp/sp/presets/all';

export interface IDynamicDropdownState{
  singleValueDropdown:string,
  multiValueDropdown:any
}

export default class DynamicDropdown extends React.Component<IDynamicDropdownProps, IDynamicDropdownState> {

  constructor(props:any){
    super(props);
    sp.setup({
      spfxContext:this.context as any
    });
    this.state = {
      singleValueDropdown:"",
      multiValueDropdown:[]
    }
  }

  private onDropdownChange=(event:React.FormEvent<HTMLDivElement>,item:IDropdownOption):void=>{
    this.setState({singleValueDropdown:item?.key as string });
  }
  private onmultiselectDropdown=(event:React.FormEvent<HTMLDivElement>,item:IDropdownOption):void=>{
    const selectedkeys=item?.selected?[...this.state.multiValueDropdown,item.key as string]:
    this.state.multiValueDropdown.filter((key:any)=>key!==item.key);
    this.setState({multiValueDropdown:selectedkeys});
  }

  public async saveData(e:any){
    const web=Web(this.props.siteUrl);
    await web.lists.getByTitle('Fluentuidropdown').items.add({
      Title:Guid.newGuid().toString(),
      singleValueDropdown:this.state.singleValueDropdown,
      multiValueDropdown:{results:this.state.multiValueDropdown}
    })
    .then((data)=>{
      console.log(" no Error found");
      return data;
    })
    .catch((err)=>{
      console.error("error found");
      throw err;
    });
    Dialog.alert("submitted sucessfully");
  }

  public render(): React.ReactElement<IDynamicDropdownProps> {
   

    return (
      <>
      <Dropdown placeholder='Single select dropdown'
      options={this.props.singleValueOptions}
      selectedKey={this.state.singleValueDropdown}
      label="Single Selected Dropdown"
      onChange={this.onDropdownChange}
      />
      <Dropdown placeholder='multi select dropdown'
      defaultSelectedKeys={this.state.multiValueDropdown}
      label='multi select dropdown'
      multiSelect
      options={this.props.multiValueOptions}
      onChange={this.onmultiselectDropdown}
      />
      <br/>
      <PrimaryButton text='Save' onClick={e=>this.saveData(e)}/>

      </>
    );
  }
}
