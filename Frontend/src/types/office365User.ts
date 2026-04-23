export interface Office365User {
  Id: string;
  DisplayName: string;
  Mail: string | null;
  UserPrincipalName: string;
  GivenName: string | null;
  Surname: string | null;
  JobTitle: string | null;
  Department: string | null;
}

export interface SearchUserResponse {
  value: Office365User[];
}
