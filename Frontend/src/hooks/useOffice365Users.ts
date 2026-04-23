import { useState, useCallback } from 'react';
import { getClient } from '@microsoft/power-apps/data';
import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo';
import type { User } from '../generated/models/Office365UsersModel';

export type { User as Office365User };

export function useOffice365Users() {
  const [users, setUsers] = useState<User[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const searchUsers = useCallback(async (searchTerm: string = '') => {
    setLoading(true);
    setError(null);
    try {
      const client = getClient(dataSourcesInfo);
      const term = searchTerm.trim() || undefined;
      const result = await client.executeAsync<
        { searchTerm?: string; top?: number; isSearchTermRequired?: boolean },
        { value?: User[] }
      >({
        connectorOperation: {
          tableName: 'office365users',
          operationName: 'SearchUserV2',
          parameters: { searchTerm: term, top: 100, isSearchTermRequired: false },
        },
      });
      if (result.success && result.data) {
        setUsers(result.data.value ?? []);
      } else {
        setError(result.error?.message ?? 'Failed to load users.');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error occurred.');
    } finally {
      setLoading(false);
    }
  }, []);

  return { users, loading, error, searchUsers };
}

