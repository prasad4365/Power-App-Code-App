import { useState, useRef, useEffect } from 'react';
import { useTestData } from './hooks/useTestData';
import { useOffice365Users } from './hooks/useOffice365Users';
import { useSharePointDocuments } from './hooks/useSharePointDocuments';
import type { Office365User } from './hooks/useOffice365Users';
import type { TestDataRecord } from './types/testData';

// ── User Dropdown ────────────────────────────────────────────────────────────
function getInitials(name: string): string {
  return name
    .split(' ')
    .filter(Boolean)
    .slice(0, 2)
    .map((n) => n[0].toUpperCase())
    .join('');
}

const AVATAR_COLORS = [
  'bg-indigo-600', 'bg-violet-600', 'bg-sky-600',
  'bg-emerald-600', 'bg-amber-600', 'bg-rose-600',
];
function avatarColor(name: string): string {
  let hash = 0;
  for (const c of name) hash = (hash * 31 + c.charCodeAt(0)) & 0xffff;
  return AVATAR_COLORS[hash % AVATAR_COLORS.length];
}

function UserDropdown() {
  const { users, loading, error, searchUsers } = useOffice365Users();
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState<Office365User | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // Load users when dropdown opens
  useEffect(() => {
    if (open) {
      searchUsers(query);
      inputRef.current?.focus();
    }
  }, [open]);

  // Debounced search as user types
  useEffect(() => {
    if (!open) return;
    const t = setTimeout(() => searchUsers(query), 300);
    return () => clearTimeout(t);
  }, [query, open]);

  // Close on outside click
  useEffect(() => {
    function handler(e: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    }
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  function handleSelect(user: Office365User) {
    setSelected(user);
    setOpen(false);
    setQuery('');
  }

  function handleClear() {
    setSelected(null);
    setQuery('');
  }

  return (
    <div className="mb-8">
      <label className="block text-xs font-medium text-slate-400 mb-2">
        <span className="flex items-center gap-1.5">
          <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z" />
          </svg>
          Select User (Office 365)
        </span>
      </label>

      <div ref={containerRef} className="relative">
        {/* Trigger / selected display */}
        <button
          type="button"
          onClick={() => setOpen((o) => !o)}
          className="w-full flex items-center justify-between gap-3 px-4 py-3 rounded-xl bg-slate-800 border border-slate-700 hover:border-slate-500 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition text-left"
        >
          {selected ? (
            <div className="flex items-center gap-3 min-w-0">
              <div className={`w-8 h-8 rounded-full ${avatarColor(selected.DisplayName ?? '')} flex items-center justify-center text-white text-xs font-semibold shrink-0`}>
                {getInitials(selected.DisplayName ?? '')}
              </div>
              <div className="min-w-0">
                <p className="text-sm font-medium text-white truncate">{selected.DisplayName}</p>
                <p className="text-xs text-slate-400 truncate">{selected.Mail ?? selected.UserPrincipalName}</p>
              </div>
            </div>
          ) : (
            <span className="text-slate-500 text-sm">Search and select a user...</span>
          )}
          <div className="flex items-center gap-1 shrink-0">
            {selected && (
              <span
                role="button"
                onClick={(e) => { e.stopPropagation(); handleClear(); }}
                className="p-1 rounded text-slate-500 hover:text-white transition"
              >
                <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </span>
            )}
            <svg className={`w-4 h-4 text-slate-500 transition-transform ${open ? 'rotate-180' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
            </svg>
          </div>
        </button>

        {/* Dropdown panel */}
        {open && (
          <div className="absolute z-20 w-full mt-2 bg-slate-800 border border-slate-700 rounded-2xl shadow-2xl overflow-hidden">
            {/* Search input */}
            <div className="p-3 border-b border-slate-700">
              <div className="relative">
                <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                </svg>
                <input
                  ref={inputRef}
                  type="text"
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  placeholder="Type a name or email..."
                  className="w-full pl-9 pr-4 py-2 rounded-lg bg-slate-900 border border-slate-600 text-sm text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition"
                />
              </div>
            </div>

            {/* List */}
            <div className="max-h-64 overflow-y-auto">
              {loading && (
                <div className="flex items-center justify-center py-8 gap-2">
                  <div className="w-5 h-5 rounded-full border-2 border-indigo-500 border-t-transparent animate-spin" />
                  <span className="text-slate-400 text-sm">Loading users...</span>
                </div>
              )}
              {!loading && error && (
                <p className="text-center py-6 text-red-400 text-sm px-4">{error}</p>
              )}
              {!loading && !error && users.length === 0 && (
                <p className="text-center py-8 text-slate-500 text-sm">No users found</p>
              )}
              {!loading && !error && users.map((user) => (
                <button
                  key={user.Id}
                  type="button"
                  onClick={() => handleSelect(user)}
                  className={`w-full flex items-center gap-3 px-4 py-3 text-left hover:bg-slate-700 transition-colors ${
                    selected?.Id === user.Id ? 'bg-indigo-900/30' : ''
                  }`}
                >
                  <div className={`w-9 h-9 rounded-full ${avatarColor(user.DisplayName ?? '')} flex items-center justify-center text-white text-xs font-semibold shrink-0`}>
                    {getInitials(user.DisplayName ?? '')}
                  </div>
                  <div className="min-w-0">
                    <p className="text-sm font-medium text-white truncate">{user.DisplayName}</p>
                    <p className="text-xs text-slate-400 truncate">{user.Mail ?? user.UserPrincipalName}</p>
                  </div>
                  {selected?.Id === user.Id && (
                    <svg className="w-4 h-4 text-indigo-400 ml-auto shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                    </svg>
                  )}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Selected user profile card */}
      {selected && (
        <div className="mt-3 p-4 bg-slate-800/60 border border-slate-700 rounded-2xl flex items-center gap-4">
          <div className={`w-12 h-12 rounded-full ${avatarColor(selected.DisplayName ?? '')} flex items-center justify-center text-white text-base font-bold shrink-0`}>
            {getInitials(selected.DisplayName ?? '')}
          </div>
          <div className="min-w-0 flex-1">
            <p className="font-semibold text-white">{selected.DisplayName}</p>
            <p className="text-xs text-slate-400 truncate">{selected.Mail ?? selected.UserPrincipalName}</p>
            {(selected.JobTitle || selected.Department) && (
              <p className="text-xs text-slate-500 mt-0.5">
                {[selected.JobTitle, selected.Department].filter(Boolean).join(' · ')}
              </p>
            )}
          </div>
          <span className="text-xs text-emerald-400 bg-emerald-900/30 border border-emerald-800/40 px-2 py-1 rounded-lg shrink-0">Selected</span>
        </div>
      )}
    </div>
  );
}

// ── Add Record Modal ─────────────────────────────────────────────────────────
function AddRecordModal({
  onClose,
  onSubmit,
  saving,
}: {
  onClose: () => void;
  onSubmit: (data: { Title: string; Description: string | null }) => Promise<boolean>;
  saving: boolean;
}) {
  const [title, setTitle] = useState('');
  const [description, setDescription] = useState('');
  const [titleError, setTitleError] = useState('');

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!title.trim()) {
      setTitleError('Title is required.');
      return;
    }
    const ok = await onSubmit({
      Title: title.trim(),
      Description: description.trim() || null,
    });
    if (ok) onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
      {/* Backdrop */}
      <div className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={onClose} />
      {/* Modal */}
      <div className="relative w-full max-w-md bg-slate-800 border border-slate-700 rounded-2xl shadow-2xl p-6">
        <div className="flex items-center justify-between mb-5">
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 rounded-lg bg-indigo-600/30 flex items-center justify-center">
              <svg className="w-4 h-4 text-indigo-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
              </svg>
            </div>
            <h2 className="text-base font-semibold text-white">Add New Record</h2>
          </div>
          <button onClick={onClose} className="text-slate-500 hover:text-white transition">
            <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs font-medium text-slate-400 mb-1.5">
              Title <span className="text-red-400">*</span>
            </label>
            <input
              type="text"
              value={title}
              onChange={(e) => { setTitle(e.target.value); setTitleError(''); }}
              placeholder="Enter title..."
              maxLength={255}
              className={`w-full px-3.5 py-2.5 rounded-xl bg-slate-900 border text-sm text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition ${titleError ? 'border-red-500' : 'border-slate-600'}`}
            />
            {titleError && <p className="mt-1 text-xs text-red-400">{titleError}</p>}
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-400 mb-1.5">
              Description <span className="text-slate-600">(optional)</span>
            </label>
            <textarea
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="Enter description..."
              maxLength={255}
              rows={3}
              className="w-full px-3.5 py-2.5 rounded-xl bg-slate-900 border border-slate-600 text-sm text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition resize-none"
            />
            <p className="mt-1 text-xs text-slate-600 text-right">{description.length}/255</p>
          </div>

          <div className="flex gap-3 pt-1">
            <button
              type="button"
              onClick={onClose}
              className="flex-1 px-4 py-2.5 rounded-xl border border-slate-600 text-slate-300 hover:bg-slate-700 text-sm font-medium transition-colors"
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={saving}
              className="flex-1 px-4 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 disabled:opacity-50 disabled:cursor-not-allowed text-white text-sm font-medium transition-colors flex items-center justify-center gap-2"
            >
              {saving ? (
                <>
                  <svg className="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                  </svg>
                  Saving...
                </>
              ) : 'Add Record'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

// ── Delete Confirm Modal ──────────────────────────────────────────────────────
function DeleteConfirmModal({
  record,
  onClose,
  onConfirm,
  saving,
}: {
  record: TestDataRecord;
  onClose: () => void;
  onConfirm: () => Promise<boolean>;
  saving: boolean;
}) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
      <div className="absolute inset-0 bg-black/60 backdrop-blur-sm" onClick={onClose} />
      <div className="relative w-full max-w-sm bg-slate-800 border border-slate-700 rounded-2xl shadow-2xl p-6">
        <div className="flex items-center gap-3 mb-4">
          <div className="w-10 h-10 rounded-xl bg-red-900/40 border border-red-800/40 flex items-center justify-center shrink-0">
            <svg className="w-5 h-5 text-red-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
            </svg>
          </div>
          <div>
            <h2 className="text-base font-semibold text-white">Delete Record</h2>
            <p className="text-xs text-slate-400">This action cannot be undone.</p>
          </div>
        </div>
        <div className="bg-slate-900/60 rounded-xl p-3 mb-5 border border-slate-700">
          <p className="text-xs text-slate-500 mb-0.5">#{record.ID}</p>
          <p className="text-sm text-white font-medium truncate">{record.Title}</p>
        </div>
        <div className="flex gap-3">
          <button
            onClick={onClose}
            className="flex-1 px-4 py-2.5 rounded-xl border border-slate-600 text-slate-300 hover:bg-slate-700 text-sm font-medium transition-colors"
          >
            Cancel
          </button>
          <button
            onClick={onConfirm}
            disabled={saving}
            className="flex-1 px-4 py-2.5 rounded-xl bg-red-700 hover:bg-red-600 disabled:opacity-50 disabled:cursor-not-allowed text-white text-sm font-medium transition-colors flex items-center justify-center gap-2"
          >
            {saving ? (
              <svg className="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
              </svg>
            ) : null}
            Delete
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Record Card ───────────────────────────────────────────────────────────────
function RecordCard({
  record,
  onDelete,
}: {
  record: TestDataRecord;
  onDelete: (record: TestDataRecord) => void;
}) {
  return (
    <div className="group bg-slate-800 border border-slate-700 rounded-2xl p-5 hover:border-indigo-500/50 hover:shadow-lg hover:shadow-indigo-900/20 transition-all duration-200">
      <div className="flex items-start justify-between gap-2 mb-3">
        <span className="inline-flex items-center px-2.5 py-1 rounded-lg bg-indigo-900/50 border border-indigo-700/50 text-indigo-300 text-xs font-mono font-semibold">
          #{record.ID}
        </span>
        <button
          onClick={() => onDelete(record)}
          className="opacity-0 group-hover:opacity-100 p-1.5 rounded-lg text-slate-500 hover:text-red-400 hover:bg-red-900/30 transition-all"
          title="Delete record"
        >
          <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
          </svg>
        </button>
      </div>
      <h3 className="text-white font-semibold text-base leading-snug mb-2 group-hover:text-indigo-300 transition-colors">
        {record.Title}
      </h3>
      {record.Description ? (
        <p className="text-slate-400 text-sm leading-relaxed line-clamp-3">{record.Description}</p>
      ) : (
        <p className="text-slate-600 text-sm italic">No description</p>
      )}
    </div>
  );
}

// ── Sidebar ───────────────────────────────────────────────────────────────────
type Page = 'dashboard' | 'records' | 'users' | 'documents';

// ── Document helpers ─────────────────────────────────────────────────────────
function fileIcon(name: string) {
  const ext = name.split('.').pop()?.toLowerCase() ?? '';
  if (['doc', 'docx'].includes(ext))
    return <span className="text-blue-400 font-bold text-xs">W</span>;
  if (['xls', 'xlsx'].includes(ext))
    return <span className="text-emerald-400 font-bold text-xs">X</span>;
  if (['ppt', 'pptx'].includes(ext))
    return <span className="text-orange-400 font-bold text-xs">P</span>;
  if (ext === 'pdf')
    return <span className="text-red-400 font-bold text-xs">PDF</span>;
  if (['png', 'jpg', 'jpeg', 'gif', 'webp'].includes(ext))
    return (
      <svg className="w-3.5 h-3.5 text-violet-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
      </svg>
    );
  return (
    <svg className="w-3.5 h-3.5 text-slate-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
    </svg>
  );
}

function formatDate(iso: string): string {
  try {
    return new Date(iso).toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' });
  } catch {
    return iso;
  }
}

function DocumentsPage() {
  const { documents, loading, error, refetch } = useSharePointDocuments();
  const [search, setSearch] = useState('');

  // Load documents when the page mounts
  useEffect(() => { refetch(); }, []);

  const filtered = documents.filter((d) =>
    (d['{FilenameWithExtension}'] ?? d.Title ?? '').toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div>
      {/* Toolbar */}
      <div className="flex items-center justify-between gap-4 mb-6">
        <div className="relative flex-1 max-w-sm">
          <svg className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
          </svg>
          <input
            type="text"
            placeholder="Search documents..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="w-full pl-10 pr-4 py-2.5 rounded-xl bg-slate-800 border border-slate-700 text-sm text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition"
          />
        </div>
        <button
          onClick={refetch}
          disabled={loading}
          className="flex items-center gap-2 px-3 py-2 rounded-xl border border-slate-700 hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed text-sm font-medium transition-colors"
        >
          <svg className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
          </svg>
          Refresh
        </button>
      </div>

      {/* Error */}
      {error && !loading && (
        <div className="flex flex-col items-center justify-center py-24 gap-4">
          <div className="w-14 h-14 rounded-2xl bg-red-900/30 border border-red-800/40 flex items-center justify-center">
            <svg className="w-7 h-7 text-red-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" />
            </svg>
          </div>
          <p className="font-semibold text-red-400">Failed to load documents</p>
          <p className="text-slate-400 text-sm">{error}</p>
          <button onClick={refetch} className="px-5 py-2.5 rounded-xl bg-red-700 hover:bg-red-600 text-sm font-medium transition-colors">Try Again</button>
        </div>
      )}

      {/* Loading */}
      {loading && (
        <div className="flex flex-col items-center justify-center py-32 gap-4">
          <div className="w-12 h-12 rounded-full border-4 border-indigo-600 border-t-transparent animate-spin" />
          <p className="text-slate-400 text-sm">Loading from SharePoint...</p>
        </div>
      )}

      {/* Document table */}
      {!loading && !error && (
        <div className="bg-slate-800/60 border border-slate-700 rounded-2xl overflow-hidden">
          {/* Table header */}
          <div className="grid grid-cols-[2fr_1fr_1fr] gap-4 px-5 py-3 border-b border-slate-700 bg-slate-900/40">
            <span className="text-xs font-medium text-slate-500 uppercase tracking-wide">Name</span>
            <span className="text-xs font-medium text-slate-500 uppercase tracking-wide">Modified</span>
            <span className="text-xs font-medium text-slate-500 uppercase tracking-wide">Modified By</span>
          </div>

          {filtered.length === 0 && (
            <div className="flex flex-col items-center justify-center py-20 gap-3">
              <svg className="w-10 h-10 text-slate-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M5 19a2 2 0 01-2-2V7a2 2 0 012-2h4l2 2h4a2 2 0 012 2v1M5 19h14a2 2 0 002-2v-5a2 2 0 00-2-2H9a2 2 0 00-2 2v5a2 2 0 01-2 2z" />
              </svg>
              <p className="text-slate-500 text-sm">
                {search ? `No documents matching "${search}"` : 'No documents found in this library'}
              </p>
            </div>
          )}

          {filtered.map((doc) => {
            const name = doc['{FilenameWithExtension}'] ?? doc.Title ?? 'Untitled';
            const SP_ORIGIN = 'https://netorgft6455485.sharepoint.com';

            // Build the best possible absolute URL by trying fields in priority order
            const makeAbsolute = (p: string | undefined) =>
              p ? (p.startsWith('http') ? p : `${SP_ORIGIN}${p.startsWith('/') ? '' : '/'}${p}`) : null;

            const link =
              makeAbsolute(doc['{Link}']) ??
              makeAbsolute(doc['{FullPath}']) ??
              (doc['{Path}'] && name ? makeAbsolute(`${doc['{Path}']}/${name}`) : null) ??
              '#';

            const modifiedBy = doc.Editor?.DisplayName ?? '';
            const modified = doc.Modified ?? '';
            return (
              <div
                key={String(doc.ID)}
                className="grid grid-cols-[2fr_1fr_1fr] gap-4 px-5 py-3.5 border-t border-slate-700/50 hover:bg-slate-700/30 transition-colors"
              >
                <div className="flex items-center gap-3 min-w-0">
                  <div className="w-7 h-7 rounded-lg bg-slate-700 border border-slate-600 flex items-center justify-center shrink-0">
                    {fileIcon(name)}
                  </div>
                  <a
                    href={link}
                    title={link}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-sm text-white truncate hover:text-indigo-300 transition-colors"
                  >
                    {name}
                  </a>
                </div>
                <span className="text-sm text-slate-400 self-center truncate">{formatDate(modified)}</span>
                <span className="text-sm text-slate-400 self-center truncate">{modifiedBy}</span>
              </div>
            );
          })}
        </div>
      )}

      {!loading && !error && filtered.length > 0 && (
        <p className="text-center text-xs text-slate-600 mt-6">
          {filtered.length} document{filtered.length !== 1 ? 's' : ''} · SharePoint Document Library
        </p>
      )}
    </div>
  );
}

function Sidebar({ active, onNavigate }: { active: Page; onNavigate: (p: Page) => void }) {
  const items: { id: Page; label: string; icon: React.ReactNode }[] = [
    {
      id: 'dashboard',
      label: 'Dashboard',
      icon: (
        <svg className="w-4 h-4 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 12l2-2m0 0l7-7 7 7M5 10v10a1 1 0 001 1h3m10-11l2 2m-2-2v10a1 1 0 01-1 1h-3m-6 0a1 1 0 001-1v-4a1 1 0 011-1h2a1 1 0 011 1v4a1 1 0 001 1m-6 0h6" />
        </svg>
      ),
    },
    {
      id: 'records',
      label: 'Records',
      icon: (
        <svg className="w-4 h-4 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 10h16M4 14h8M4 18h8" />
        </svg>
      ),
    },
    {
      id: 'users',
      label: 'Users',
      icon: (
        <svg className="w-4 h-4 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z" />
        </svg>
      ),
    },
    {
      id: 'documents',
      label: 'Documents',
      icon: (
        <svg className="w-4 h-4 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 19a2 2 0 01-2-2V7a2 2 0 012-2h4l2 2h4a2 2 0 012 2v1M5 19h14a2 2 0 002-2v-5a2 2 0 00-2-2H9a2 2 0 00-2 2v5a2 2 0 01-2 2z" />
        </svg>
      ),
    },
  ];

  return (
    <aside className="fixed top-0 left-0 h-screen w-60 bg-slate-900 border-r border-slate-700/60 flex flex-col z-20">
      {/* Brand */}
      <div className="px-5 py-5 border-b border-slate-700/60 flex items-center gap-3">
        <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-indigo-600 to-violet-600 flex items-center justify-center shadow-md shrink-0">
          <svg className="w-4 h-4 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 10h16M4 14h8M4 18h8" />
          </svg>
        </div>
        <div className="min-w-0">
          <p className="text-sm font-semibold text-white truncate">Test Data App</p>
          <p className="text-xs text-slate-500 truncate">Power Apps Code App</p>
        </div>
      </div>

      {/* Nav items */}
      <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto">
        {items.map((item) => (
          <button
            key={item.id}
            onClick={() => onNavigate(item.id)}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium transition-all ${
              active === item.id
                ? 'bg-indigo-600/20 text-indigo-300 border border-indigo-700/40'
                : 'text-slate-400 hover:text-white hover:bg-slate-800 border border-transparent'
            }`}
          >
            {item.icon}
            {item.label}
          </button>
        ))}
      </nav>

      {/* Footer */}
      <div className="px-5 py-4 border-t border-slate-700/60">
        <p className="text-xs text-slate-600 truncate">dbo.TestData · SQL Server</p>
      </div>
    </aside>
  );
}

// ── App ───────────────────────────────────────────────────────────────────────
function App() {
  const { records, loading, saving, error, refetch, addRecord, deleteRecord } = useTestData();
  const [activePage, setActivePage] = useState<Page>('dashboard');
  const [search, setSearch] = useState('');
  const [showAddModal, setShowAddModal] = useState(false);
  const [recordToDelete, setRecordToDelete] = useState<TestDataRecord | null>(null);

  const filtered = records.filter(
    (r) =>
      r.Title.toLowerCase().includes(search.toLowerCase()) ||
      (r.Description ?? '').toLowerCase().includes(search.toLowerCase())
  );

  const pageMeta: Record<Page, { title: string; subtitle: string }> = {
    dashboard: { title: 'Dashboard', subtitle: 'Overview of your data' },
    records: { title: 'Records', subtitle: 'dbo.TestData · SQL Server' },
    users: { title: 'Users', subtitle: 'Office 365 Directory' },
    documents: { title: 'Documents', subtitle: 'SharePoint Document Library' },
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 text-white">

      {/* Sidebar */}
      <Sidebar active={activePage} onNavigate={setActivePage} />

      {/* Modals */}
      {showAddModal && (
        <AddRecordModal
          onClose={() => setShowAddModal(false)}
          onSubmit={addRecord}
          saving={saving}
        />
      )}
      {recordToDelete && (
        <DeleteConfirmModal
          record={recordToDelete}
          onClose={() => setRecordToDelete(null)}
          onConfirm={async () => {
            const ok = await deleteRecord(recordToDelete.ID);
            if (ok) setRecordToDelete(null);
            return ok;
          }}
          saving={saving}
        />
      )}

      {/* Content area — offset by sidebar width */}
      <div className="ml-60 flex flex-col min-h-screen">

        {/* Header */}
        <header className="border-b border-slate-700/60 bg-slate-900/80 backdrop-blur-sm sticky top-0 z-10">
          <div className="px-6 py-4 flex items-center justify-between">
            <div>
              <h1 className="text-base font-semibold">{pageMeta[activePage].title}</h1>
              <p className="text-xs text-slate-400">{pageMeta[activePage].subtitle}</p>
            </div>
            {activePage === 'records' && (
              <div className="flex items-center gap-2">
                <button
                  onClick={refetch}
                  disabled={loading}
                  className="flex items-center gap-2 px-3 py-2 rounded-xl border border-slate-700 hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed text-sm font-medium transition-colors"
                >
                  <svg className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                  </svg>
                  Refresh
                </button>
                <button
                  onClick={() => setShowAddModal(true)}
                  className="flex items-center gap-2 px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-sm font-medium transition-colors shadow"
                >
                  <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
                  </svg>
                  Add Record
                </button>
              </div>
            )}
          </div>
        </header>

        <main className="flex-1 px-6 py-8">

          {/* ── Dashboard ── */}
          {activePage === 'dashboard' && (
            <div>
              {loading && (
                <div className="flex flex-col items-center justify-center py-32 gap-4">
                  <div className="w-12 h-12 rounded-full border-4 border-indigo-600 border-t-transparent animate-spin" />
                  <p className="text-slate-400 text-sm">Loading...</p>
                </div>
              )}
              {!loading && !error && (
                <>
                  {/* Stats */}
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-4 mb-8">
                    <div className="bg-slate-800/60 border border-slate-700 rounded-2xl px-5 py-4">
                      <p className="text-3xl font-bold text-white">{records.length}</p>
                      <p className="text-xs text-slate-400 mt-1">Total Records</p>
                    </div>
                    <div className="bg-slate-800/60 border border-slate-700 rounded-2xl px-5 py-4">
                      <p className="text-3xl font-bold text-emerald-400">
                        {records.filter((r) => r.Description).length}
                      </p>
                      <p className="text-xs text-slate-400 mt-1">With Description</p>
                    </div>
                    <div className="bg-slate-800/60 border border-slate-700 rounded-2xl px-5 py-4">
                      <p className="text-3xl font-bold text-slate-400">
                        {records.filter((r) => !r.Description).length}
                      </p>
                      <p className="text-xs text-slate-400 mt-1">No Description</p>
                    </div>
                  </div>

                  {/* Recent records */}
                  {records.length > 0 ? (
                    <div>
                      <div className="flex items-center justify-between mb-4">
                        <h2 className="text-sm font-semibold text-slate-300">Recent Records</h2>
                        <button
                          onClick={() => setActivePage('records')}
                          className="text-xs text-indigo-400 hover:text-indigo-300 transition-colors"
                        >
                          View all →
                        </button>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                        {records.slice(0, 6).map((record) => (
                          <RecordCard key={record.ID} record={record} onDelete={setRecordToDelete} />
                        ))}
                      </div>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center justify-center py-24 gap-4">
                      <div className="w-14 h-14 rounded-2xl bg-slate-700/50 flex items-center justify-center">
                        <svg className="w-7 h-7 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M20 13V6a2 2 0 00-2-2H6a2 2 0 00-2 2v7m16 0v5a2 2 0 01-2 2H6a2 2 0 01-2-2v-5m16 0h-2.586a1 1 0 00-.707.293l-2.414 2.414a1 1 0 01-.707.293h-3.172a1 1 0 01-.707-.293l-2.414-2.414A1 1 0 006.586 13H4" />
                        </svg>
                      </div>
                      <p className="text-slate-400 text-sm">No records yet. Go to Records to add some.</p>
                    </div>
                  )}
                </>
              )}
              {error && !loading && (
                <div className="flex flex-col items-center justify-center py-24 gap-4">
                  <p className="font-semibold text-red-400">Failed to load data</p>
                  <p className="text-slate-400 text-sm">{error}</p>
                  <button onClick={refetch} className="px-5 py-2.5 rounded-xl bg-red-700 hover:bg-red-600 text-sm font-medium transition-colors">Try Again</button>
                </div>
              )}
            </div>
          )}

          {/* ── Records ── */}
          {activePage === 'records' && (
            <div>
              {/* Search */}
              {!loading && !error && records.length > 0 && (
                <div className="relative mb-6">
                  <svg className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                  </svg>
                  <input
                    type="text"
                    placeholder="Search by title or description..."
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    className="w-full pl-10 pr-10 py-3 rounded-xl bg-slate-800 border border-slate-700 text-sm text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent transition"
                  />
                  {search && (
                    <button onClick={() => setSearch('')} className="absolute right-3.5 top-1/2 -translate-y-1/2 text-slate-500 hover:text-white transition">
                      <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                      </svg>
                    </button>
                  )}
                </div>
              )}

              {loading && (
                <div className="flex flex-col items-center justify-center py-32 gap-4">
                  <div className="w-12 h-12 rounded-full border-4 border-indigo-600 border-t-transparent animate-spin" />
                  <p className="text-slate-400 text-sm">Loading from SQL Server...</p>
                </div>
              )}
              {error && !loading && (
                <div className="flex flex-col items-center justify-center py-24 gap-4">
                  <div className="w-14 h-14 rounded-2xl bg-red-900/30 border border-red-800/40 flex items-center justify-center">
                    <svg className="w-7 h-7 text-red-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" />
                    </svg>
                  </div>
                  <div className="text-center max-w-md">
                    <p className="font-semibold text-red-400 mb-1">Failed to load data</p>
                    <p className="text-slate-400 text-sm">{error}</p>
                  </div>
                  <button onClick={refetch} className="px-5 py-2.5 rounded-xl bg-red-700 hover:bg-red-600 text-sm font-medium transition-colors">Try Again</button>
                </div>
              )}
              {!loading && !error && records.length === 0 && (
                <div className="flex flex-col items-center justify-center py-24 gap-4">
                  <div className="w-14 h-14 rounded-2xl bg-slate-700/50 flex items-center justify-center">
                    <svg className="w-7 h-7 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M20 13V6a2 2 0 00-2-2H6a2 2 0 00-2 2v7m16 0v5a2 2 0 01-2 2H6a2 2 0 01-2-2v-5m16 0h-2.586a1 1 0 00-.707.293l-2.414 2.414a1 1 0 01-.707.293h-3.172a1 1 0 01-.707-.293l-2.414-2.414A1 1 0 006.586 13H4" />
                    </svg>
                  </div>
                  <div className="text-center">
                    <p className="text-slate-400 text-sm mb-3">No records in dbo.TestData</p>
                    <button onClick={() => setShowAddModal(true)} className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-sm font-medium transition-colors">
                      Add First Record
                    </button>
                  </div>
                </div>
              )}
              {!loading && !error && records.length > 0 && filtered.length === 0 && (
                <div className="text-center py-16">
                  <p className="text-slate-400 text-sm">No records match "<span className="text-slate-300">{search}</span>"</p>
                  <button onClick={() => setSearch('')} className="mt-3 text-indigo-400 hover:text-indigo-300 text-sm transition-colors">Clear search</button>
                </div>
              )}
              {!loading && !error && filtered.length > 0 && (
                <>
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                    {filtered.map((record) => (
                      <RecordCard key={record.ID} record={record} onDelete={setRecordToDelete} />
                    ))}
                  </div>
                  <p className="text-center text-xs text-slate-600 mt-8">
                    Showing {filtered.length} of {records.length} records · dbo.TestData
                  </p>
                </>
              )}
            </div>
          )}

          {/* ── Users ── */}
          {activePage === 'users' && (
            <div className="max-w-xl">
              <UserDropdown />
            </div>
          )}

          {/* ── Documents ── */}
          {activePage === 'documents' && <DocumentsPage />}

        </main>
      </div>
    </div>
  );
}

export default App;
