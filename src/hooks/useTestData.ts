import { useState, useEffect, useCallback } from 'react';
import { getClient } from '@microsoft/power-apps/data';
import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo';
import type { TestDataRecord } from '../types/testData';

const DATA_SOURCE_NAME = 'testdata';

// Use the full generated dataSourcesInfo (which includes office365users) so the
// singleton PowerDataSourcesInfoProvider is initialized with ALL data sources.
// Override testdata.version to 'v2' which is required for the SQL connector.
const DATA_SOURCES_INFO = {
  ...dataSourcesInfo,
  testdata: {
    ...dataSourcesInfo.testdata,
    version: 'v2',
  },
};

export function useTestData() {
  const [records, setRecords] = useState<TestDataRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const fetchRecords = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const client = getClient(DATA_SOURCES_INFO);
      const result = await client.retrieveMultipleRecordsAsync<TestDataRecord>(DATA_SOURCE_NAME);
      if (result.success && result.data) {
        setRecords(result.data);
      } else {
        setError(result.error?.message ?? 'Failed to load records.');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error occurred.');
    } finally {
      setLoading(false);
    }
  }, []);

  const addRecord = useCallback(async (input: { Title: string; Description: string | null }): Promise<boolean> => {
    setSaving(true);
    try {
      const client = getClient(DATA_SOURCES_INFO);
      const result = await client.createRecordAsync<typeof input, TestDataRecord>(DATA_SOURCE_NAME, input);
      if (result.success) {
        await fetchRecords();
        return true;
      } else {
        setError(result.error?.message ?? 'Failed to create record.');
        return false;
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error occurred.');
      return false;
    } finally {
      setSaving(false);
    }
  }, [fetchRecords]);

  const deleteRecord = useCallback(async (id: number): Promise<boolean> => {
    setSaving(true);
    try {
      const client = getClient(DATA_SOURCES_INFO);
      const result = await client.deleteRecordAsync(DATA_SOURCE_NAME, String(id));
      if (result.success) {
        setRecords((prev) => prev.filter((r) => r.ID !== id));
        return true;
      } else {
        setError(result.error?.message ?? 'Failed to delete record.');
        return false;
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error occurred.');
      return false;
    } finally {
      setSaving(false);
    }
  }, []);

  useEffect(() => {
    fetchRecords();
  }, [fetchRecords]);

  return { records, loading, saving, error, refetch: fetchRecords, addRecord, deleteRecord };
}
