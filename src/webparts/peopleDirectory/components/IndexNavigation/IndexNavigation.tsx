import * as React from 'react';
import styles from './IndexNavigation.module.scss';
import { IIndexNavigationProps } from '.';
import { Search } from '../Search';
import { Filter } from '../Filter/Filter';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import * as strings from 'PeopleDirectoryWebPartStrings';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export class IndexNavigation extends React.Component<IIndexNavigationProps, {}> {

  private NameOptions: IDropdownOption[] = [{ key: 'FirstName', text: 'First Name' }, { key: 'LastName', text: 'Last Name' }]
  /**
   * Event handler for selecting a tab in the navigation
   */
  private _handleIndexSelect = (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>): void => {
    this.props.onIndexSelect(item.props.linkText);
  }

  private _handleDropdownChange = (event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void => {
    this.props.onDropdownNameChange(option);
  }

  // private _handleFilterChange = (departments: string[]): void => {
  //   this.props.onDepartmentFilterChange(departments);
  // }



  public shouldComponentUpdate(nextProps: IIndexNavigationProps, nextState: {}, nextContext: any): boolean {
    // Component should update only if the selected tab has changed.
    // This check helps to avoid unnecessary renders
    return this.props.selectedIndex !== nextProps.selectedIndex;
  }

  public render(): React.ReactElement<IIndexNavigationProps> {
    // build the list of alphabet letters A..Z    
    let az = Array.apply(null, { length: 26 }).map((x: string, i: number): string => { return String.fromCharCode(65 + i); });
    az.unshift('See All');
    az.push(' ');
    if (this.props.locale === "sv-SE") {
      az.push('Å', 'Ä', 'Ö');
    }
    // for each letter, create a PivotItem component
    const indexes: JSX.Element[] = az.map(index => <PivotItem linkText={index} itemKey={index} key={index} />);
    // as the last tab in the navigation, add the Search option
    // indexes.push(<PivotItem linkText={strings.SearchButtonText} itemKey='Search'>    
    //   <Search
    //     searchQuery={this.props.searchQuery}
    //     onSearch={this.props.onSearch}
    //     onClear={this.props.onSearchClear} />
    // </PivotItem>);

    return (
      <div className={styles.indexNavigation} >
        <span className='inline-block'>
          <Dropdown
            defaultSelectedKey={this.NameOptions[0].key}
            options={this.NameOptions}
            onChange={this._handleDropdownChange}
          />
        </span>
        <span className='search'>
          <Search
            searchQuery={this.props.searchQuery}
            onSearch={this.props.onSearch}
            onClear={this.props.onSearchClear} />
        </span>
        <Pivot onLinkClick={this._handleIndexSelect} selectedKey={this.props.selectedIndex}>
          {indexes}
        </Pivot>
      </div>
    );
  }
}
