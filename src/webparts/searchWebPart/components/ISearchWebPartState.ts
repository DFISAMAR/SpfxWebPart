export interface ISearchWebPartState {
  searchQuery: string;
  searchResults: { DisplayName: string; WorkEmail: string; Title: string }[]; // Adjust as necessary
}
