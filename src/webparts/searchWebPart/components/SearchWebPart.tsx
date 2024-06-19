import * as React from 'react';
import styles from './SearchWebPart.module.scss';
import { ISearchWebPartProps } from './ISearchWebPartProps';
import { sp } from '@pnp/sp/presets/all';

export interface ISearchWebPartState {
  searchQuery: string;
  allResults: { AccountName: string; WorkEmail: string; Title: string }[];
  searchResults: { AccountName: string; WorkEmail: string; Title: string }[];
}

export default class SearchWebPart extends React.Component<ISearchWebPartProps, ISearchWebPartState> {
  constructor(props: ISearchWebPartProps) {
    super(props);
    this.state = {
      searchQuery: '',
      allResults: [],
      searchResults: []
    };
  }

  public async componentDidMount(): Promise<void> {
    sp.setup({
      spfxContext: this.props.context as any // Cast to any to bypass type-checking
    });

    try {
      const allNames = await sp.web.lists.getByTitle('EmployeeDirectory').items.get();
      this.setState({ 
        allResults: allNames,
        searchResults: allNames // Initially display all results
      });
    } catch (error) {
      console.error(error);
    }
  }

  private handleInputChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const searchQuery = event.target.value;
    const searchResults = this.state.allResults.filter(item => item.AccountName.toLowerCase().includes(searchQuery.toLowerCase()));
    this.setState({ searchQuery, searchResults });
  }

  public render(): React.ReactElement<ISearchWebPartProps> {
    return (
      <div className={styles.searchWebPart}>
        <div className={styles.header}>
          <h2>Employee Directory</h2>
          <input 
            className={styles.searchBox}
            type="text" 
            value={this.state.searchQuery} 
            onChange={this.handleInputChange} 
            placeholder="Search names..."
          />
        </div>
        <ul className={styles.searchResults}>
          {this.state.searchResults.map((result, index) => (
            <li key={index} className={styles.searchResult}>
              <img src="https://trailorg255.sharepoint.com/sites/IntranetPortal/SiteAssets/Images/DummyImage.jpg" alt="Avatar" width="50" height="50"  />
              <div className={styles.details}>
                <div className={styles.name}>{result.AccountName}</div>
                <div className={styles.email}>Email: {result.WorkEmail}</div>
                <div className={styles.role}>Role: {result.Title}</div>
              </div>
            </li>
          ))}
        </ul>
      </div>
    );
  }
}
