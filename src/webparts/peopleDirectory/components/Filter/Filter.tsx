import * as React from 'react';
import { Checkbox, ICheckboxProps } from 'office-ui-fabric-react/lib/Checkbox';
// import styles from './Search.module.scss';
// import { ISearchProps } from '.';

export interface IFilterProps {
    // Departments: ICheckboxProps[]
    onFilterChange: (selectedDepartments: string[]) => void;
}

export class Filter extends React.Component<IFilterProps, {}> {
    // private _handleSearch = (searchQuery: string): void => {
    //     this.props.onSearch(searchQuery);
    // }

    // private _handleClear = (): void => {
    //     this.props.onClear();
    // }

    private departments = ['IT', 'HR'];
    private selectedDeaprtments: string[] = [];

    private _handleFilterChange = (index: number, ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean): void => {
        if (isChecked) {
            this.selectedDeaprtments.push(this.departments[index]);
        } else {
            this.selectedDeaprtments = this.selectedDeaprtments.filter(dept => {
                return dept !== this.departments[index];
            });
        }
        this.props.onFilterChange(this.selectedDeaprtments);
    }

    public render(): React.ReactElement<IFilterProps> {
        const departmentNHTML: JSX.Element[] = this.departments.map((dept: any, index: number) => {
            return (
                <ul>
                    <Checkbox
                        label={dept}
                        onChange={this._handleFilterChange.bind(this, index)}
                    />
                </ul>
            )
        });

        return (
            <div >
                <h3 style={{ fontSize: '20px', fontFamily: 'serif' }}>Departments</h3>
                <br />
                {
                    departmentNHTML
                }
            </div>
        );
    }
}
