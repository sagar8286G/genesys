import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
/**
 * Properties for the index navigation component
 */
export interface IIndexNavigationProps {
  /**
   * Name of the currently selected tab, eg. 'A'
   */
  selectedIndex: string;
  /**
   * Current search query. Empty, if not searching.
   */
  searchQuery: string;
  /**
 * Current locale
 */
  locale: string;
  /**
    * Event handler for selecting a tab
    */
  // dropdownName: IDropdownOption;

  onIndexSelect: (index: string) => void;
  /**
   * Event handler for issuing a search query
   */
  onSearch: (searchQuery: string) => void;
  /**
   * Event handler for clearing the search query
   */
  onSearchClear: () => void;

  onDropdownNameChange: (dropdownOption: IDropdownOption) => void;

  // onDepartmentFilterChange: (dropdownOption: string[]) => void;

}
