import { useState, useCallback } from 'react';
import { getClient } from '@microsoft/power-apps/data';
import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo';
import type { DocumentsRead } from '../generated/models/DocumentsModel';

export type { DocumentsRead as SharePointDocument };

export function useSharePointDocuments() {
  const [documents, setDocuments] = useState<DocumentsRead[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const connected = true;

  const refetch = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const client = getClient({
        ...dataSourcesInfo,
        testdata: { ...dataSourcesInfo.testdata, version: 'v2' },
      });
      const result = await client.retrieveMultipleRecordsAsync<DocumentsRead>('documents');
      if (result.success && result.data) {
        const files = result.data.filter((d) => !d['{IsFolder}']);
        // Log raw data so URL fields can be inspected in browser devtools
        console.log('[SharePoint] raw documents:', files);
        setDocuments(files);
      } else {
        setError(result.error?.message ?? 'Failed to load documents.');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error occurred.');
    } finally {
      setLoading(false);
    }
  }, []);

  return { documents, loading, error, connected, refetch };
}

