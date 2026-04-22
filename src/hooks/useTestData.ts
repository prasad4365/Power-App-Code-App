import { useState, useEffect, useCallback } from 'react';
import type { TestDataRecord } from '../types/testData';

// Points to the Express backend.
// Locally: set VITE_API_BASE_URL=http://localhost:3001 in .env.local
// On Azure: set VITE_API_BASE_URL to your App Service URL before running npm run build
const API_BASE = import.meta.env.VITE_API_BASE_URL as string;

export function useTestData() {
  const [records, setRecords] = useState<TestDataRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const fetchRecords = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`${API_BASE}/api/testdata`);
      if (!res.ok) throw new Error(`Server responded with ${res.status}`);
      const data: TestDataRecord[] = await res.json();
      setRecords(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to load records.');
    } finally {
      setLoading(false);
    }
  }, []);

  const addRecord = useCallback(async (input: { Title: string; Description: string | null }): Promise<boolean> => {
    setSaving(true);
    try {
      const res = await fetch(`${API_BASE}/api/testdata`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(input),
      });
      if (!res.ok) throw new Error(`Server responded with ${res.status}`);
      await fetchRecords();
      return true;
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to create record.');
      return false;
    } finally {
      setSaving(false);
    }
  }, [fetchRecords]);

  const deleteRecord = useCallback(async (id: number): Promise<boolean> => {
    setSaving(true);
    try {
      const res = await fetch(`${API_BASE}/api/testdata/${id}`, {
        method: 'DELETE',
      });
      if (!res.ok) throw new Error(`Server responded with ${res.status}`);
      setRecords((prev) => prev.filter((r) => r.ID !== id));
      return true;
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to delete record.');
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
