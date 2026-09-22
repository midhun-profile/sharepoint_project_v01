import { useState, useEffect, useMemo } from 'react';
import {
  X,
  Plus,
  Trash2,
  Check,
  Search,
  Globe,
  Database,
  CheckCheck,
  RotateCcw,
  Sparkles,
  Info,
  Layers,
  ChevronDown,
  ChevronRight,
  ExternalLink,
} from 'lucide-react';
import { useConnectionsStore } from '../../state/connectionsStore';
import { DEMO_SHAREPOINT_LISTS } from '../../data/demoSharePointLists';
import { SHAREPOINT_SITE_PRESETS } from '../../data/defaultConnections';
import { AVAILABLE_ICONS, COLOR_VARIANTS, SharePointIcon, type ConnectionColor } from './SharePointIcon';
import type { ConnectionFormData, SharePointSourceConfig } from '../../types/connections';
import { detectSharePointFieldType } from '../table/fieldRenderers';

export function ConnectionConfigModal() {
  const {
    modalState,
    closeModal,
    addConnection,
    updateConnection,
    deleteConnection,
    connections,
  } = useConnectionsStore();

  const isOpen = modalState.isOpen;
  const isEditing = modalState.mode === 'edit';
  const editingId = modalState.connectionId;

  // Find target connection if in edit mode
  const existingConnection = useMemo(() => {
    if (!isEditing || !editingId) return null;
    return connections.find((c) => c.id === editingId) || null;
  }, [isEditing, editingId, connections]);

  // Form states
  const [displayName, setDisplayName] = useState('');
  const [description, setDescription] = useState('');
  const [selectedIcon, setSelectedIcon] = useState('Layers');
  const [selectedColor, setSelectedColor] = useState<ConnectionColor>('blue');
  const [sources, setSources] = useState<SharePointSourceConfig[]>([]);
  const [activeSourceIndex, setActiveSourceIndex] = useState<number>(0);
  const [columnSearchQueries, setColumnSearchQueries] = useState<Record<string, string>>({});
  const [formErrors, setFormErrors] = useState<{ displayName?: string; sources?: string }>({});
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);

  // Initialize form state when opening
  useEffect(() => {
    if (!isOpen) {
      setShowDeleteConfirm(false);
      setFormErrors({});
      return;
    }

    if (existingConnection) {
      setDisplayName(existingConnection.displayName);
      setDescription(existingConnection.description || '');
      setSelectedIcon(existingConnection.icon || 'Layers');
      setSelectedColor(existingConnection.color || 'blue');

      // Populate sources, ensuring at least one
      if (existingConnection.sources && existingConnection.sources.length > 0) {
        setSources(JSON.parse(JSON.stringify(existingConnection.sources)));
      } else {
        // Fallback for single-list legacy connection
        const defSite = SHAREPOINT_SITE_PRESETS[0];
        const defList = DEMO_SHAREPOINT_LISTS[0];
        setSources([
          {
            id: `src-${Date.now()}-0`,
            siteId: defSite.id,
            siteName: defSite.name,
            siteUrl: defSite.url,
            listId: defList.id,
            listName: defList.displayName,
            selectedColumnNames: defList.columns.map((c) => c.name),
          },
        ]);
      }
      setActiveSourceIndex(0);
    } else {
      // Create new multi-source menu connection default
      const site1 = SHAREPOINT_SITE_PRESETS[0];
      const list1 = DEMO_SHAREPOINT_LISTS[0];
      const site2 = SHAREPOINT_SITE_PRESETS[1];
      const list2 = DEMO_SHAREPOINT_LISTS[1];

      setDisplayName('Executive Projects & Hardware 360');
      setDescription('Combined view connecting multiple SharePoint lists in a single table');
      setSelectedIcon('Layers');
      setSelectedColor('blue');

      setSources([
        {
          id: `src-${Date.now()}-0`,
          siteId: site1.id,
          siteName: site1.name,
          siteUrl: site1.url,
          listId: list1.id,
          listName: list1.displayName,
          selectedColumnNames: [
            'Title',
            'ArchitectureDiagram',
            'ProjectStatus',
            'RAGStatus',
            'ExecutiveSponsor',
            'ProjectBudget',
          ],
        },
        {
          id: `src-${Date.now()}-1`,
          siteId: site2.id,
          siteName: site2.name,
          siteUrl: site2.url,
          listId: list2.id,
          listName: list2.displayName,
          selectedColumnNames: [
            'DevicePhoto',
            'AssetTag',
            'DeviceModel',
            'Category',
            'AssignedTechnician',
            'PurchaseCost',
          ],
        },
      ]);
      setActiveSourceIndex(0);
    }
  }, [isOpen, existingConnection]);

  // Add another SharePoint list source to this menu
  const handleAddSource = () => {
    // Pick the next available preset site/list that isn't already used, if possible
    const usedListIds = new Set(sources.map((s) => s.listId));
    const nextList =
      DEMO_SHAREPOINT_LISTS.find((l) => !usedListIds.has(l.id)) || DEMO_SHAREPOINT_LISTS[0];
    const nextSite =
      SHAREPOINT_SITE_PRESETS.find((s) =>
        nextList.id === 'list-project-portfolio'
          ? s.id === 'site-pmo'
          : nextList.id === 'list-it-assets'
          ? s.id === 'site-itops'
          : s.id === 'site-people'
      ) || SHAREPOINT_SITE_PRESETS[0];

    const newSource: SharePointSourceConfig = {
      id: `src-${Date.now()}-${sources.length}`,
      siteId: nextSite.id,
      siteName: nextSite.name,
      siteUrl: nextSite.url,
      listId: nextList.id,
      listName: nextList.displayName,
      selectedColumnNames: nextList.columns.slice(0, 6).map((c) => c.name),
    };

    setSources([...sources, newSource]);
    setActiveSourceIndex(sources.length);
    setFormErrors((prev) => ({ ...prev, sources: undefined }));
  };

  // Remove a source from the menu
  const handleRemoveSource = (indexToRemove: number) => {
    if (sources.length <= 1) {
      setFormErrors((prev) => ({
        ...prev,
        sources: 'A menu must connect to at least one SharePoint list.',
      }));
      return;
    }
    const updated = sources.filter((_, idx) => idx !== indexToRemove);
    setSources(updated);
    setActiveSourceIndex(Math.max(0, Math.min(activeSourceIndex, updated.length - 1)));
  };

  // Update source site
  const handleSourceSiteChange = (sourceIndex: number, siteId: string) => {
    const sitePreset = SHAREPOINT_SITE_PRESETS.find((s) => s.id === siteId);
    if (!sitePreset) return;

    setSources((prev) =>
      prev.map((s, idx) => {
        if (idx !== sourceIndex) return s;
        return {
          ...s,
          siteId: sitePreset.id,
          siteName: sitePreset.name,
          siteUrl: sitePreset.url,
        };
      })
    );
  };

  // Update source list
  const handleSourceListChange = (sourceIndex: number, listId: string) => {
    const listDef = DEMO_SHAREPOINT_LISTS.find((l) => l.id === listId);
    if (!listDef) return;

    setSources((prev) =>
      prev.map((s, idx) => {
        if (idx !== sourceIndex) return s;
        // Default to all columns of the new list
        return {
          ...s,
          listId: listDef.id,
          listName: listDef.displayName,
          selectedColumnNames: listDef.columns.slice(0, 7).map((c) => c.name),
        };
      })
    );
  };

  // Toggle column selection for a specific source
  const handleToggleColumn = (sourceIndex: number, colName: string) => {
    setSources((prev) =>
      prev.map((s, idx) => {
        if (idx !== sourceIndex) return s;
        const currentSelected = s.selectedColumnNames || [];
        const isSelected = currentSelected.includes(colName);
        let updatedCols: string[];
        if (isSelected) {
          updatedCols = currentSelected.filter((name) => name !== colName);
        } else {
          updatedCols = [...currentSelected, colName];
        }
        return {
          ...s,
          selectedColumnNames: updatedCols,
        };
      })
    );
    setFormErrors((prev) => ({ ...prev, sources: undefined }));
  };

  // Select all columns for a source
  const handleSelectAllColumns = (sourceIndex: number) => {
    const targetSource = sources[sourceIndex];
    if (!targetSource) return;
    const listDef =
      DEMO_SHAREPOINT_LISTS.find((l) => l.id === targetSource.listId) || DEMO_SHAREPOINT_LISTS[0];

    setSources((prev) =>
      prev.map((s, idx) => {
        if (idx !== sourceIndex) return s;
        return {
          ...s,
          selectedColumnNames: listDef.columns.map((c) => c.name),
        };
      })
    );
  };

  // Select minimal columns (first 5)
  const handleSelectQuickColumns = (sourceIndex: number) => {
    const targetSource = sources[sourceIndex];
    if (!targetSource) return;
    const listDef =
      DEMO_SHAREPOINT_LISTS.find((l) => l.id === targetSource.listId) || DEMO_SHAREPOINT_LISTS[0];

    setSources((prev) =>
      prev.map((s, idx) => {
        if (idx !== sourceIndex) return s;
        return {
          ...s,
          selectedColumnNames: listDef.columns.slice(0, 5).map((c) => c.name),
        };
      })
    );
  };

  // Deselect all columns for a source
  const handleDeselectAllColumns = (sourceIndex: number) => {
    setSources((prev) =>
      prev.map((s, idx) => {
        if (idx !== sourceIndex) return s;
        return {
          ...s,
          selectedColumnNames: [],
        };
      })
    );
  };

  // Calculate total columns across all sources
  const totalSelectedColumnsAcrossSources = useMemo(() => {
    return sources.reduce((sum, s) => sum + (s.selectedColumnNames?.length || 0), 0);
  }, [sources]);

  // Handle Save
  const handleSave = () => {
    const errors: { displayName?: string; sources?: string } = {};

    if (!displayName.trim()) {
      errors.displayName = 'Please enter a menu display name.';
    }

    if (sources.length === 0) {
      errors.sources = 'Please connect at least one SharePoint list.';
    } else if (totalSelectedColumnsAcrossSources === 0) {
      errors.sources = 'Please select at least one field across the connected SharePoint lists.';
    }

    if (Object.keys(errors).length > 0) {
      setFormErrors(errors);
      return;
    }

    const formData: ConnectionFormData = {
      displayName: displayName.trim(),
      description: description.trim(),
      icon: selectedIcon,
      color: selectedColor,
      sources,
    };

    if (isEditing && editingId) {
      updateConnection(editingId, formData);
    } else {
      addConnection(formData);
    }
  };

  // Handle Delete
  const handleDelete = () => {
    if (editingId) {
      deleteConnection(editingId);
      closeModal();
    }
  };

  if (!isOpen) return null;

  const currentActiveSource = sources[activeSourceIndex] || sources[0];
  const currentActiveListDef = currentActiveSource
    ? DEMO_SHAREPOINT_LISTS.find((l) => l.id === currentActiveSource.listId) ||
      DEMO_SHAREPOINT_LISTS[0]
    : DEMO_SHAREPOINT_LISTS[0];

  const currentSearchQuery = (columnSearchQueries[currentActiveSource?.id] || '').toLowerCase();
  const filteredActiveColumns = currentActiveListDef.columns.filter((c) => {
    if (!currentSearchQuery) return true;
    return (
      c.displayName.toLowerCase().includes(currentSearchQuery) ||
      c.name.toLowerCase().includes(currentSearchQuery) ||
      detectSharePointFieldType(c).toLowerCase().includes(currentSearchQuery)
    );
  });

  return (
    <div
      id="connection-config-modal-backdrop"
      className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-5 bg-neutral-900/60 backdrop-blur-xs overflow-y-auto animate-in fade-in duration-150"
    >
      <div
        id="connection-config-modal-container"
        className="bg-white rounded-2xl shadow-2xl border border-neutral-200/80 w-full max-w-4xl max-h-[92vh] flex flex-col overflow-hidden text-neutral-900"
        role="dialog"
        aria-modal="true"
        aria-labelledby="modal-title"
      >
        {/* 1. Modal Top Header */}
        <div className="flex items-center justify-between px-6 py-3.5 border-b border-neutral-200 bg-neutral-50/60 flex-shrink-0">
          <div className="flex items-center gap-3">
            <div className="p-2 rounded-xl bg-blue-100 text-blue-700">
              <SharePointIcon name={selectedIcon} color={selectedColor} className="w-5 h-5" />
            </div>
            <div>
              <h2 id="modal-title" className="text-base font-semibold text-neutral-900">
                {isEditing ? 'Configure Menu & Connected SharePoint Lists' : 'Create Menu & Connect SharePoint Lists'}
              </h2>
              <p className="text-xs text-neutral-500">
                Connect multiple SharePoint lists and select fields from each to display in a single consolidated table
              </p>
            </div>
          </div>
          <button
            id="btn-close-modal"
            type="button"
            onClick={closeModal}
            className="p-1.5 rounded-lg text-neutral-400 hover:text-neutral-700 hover:bg-neutral-200/60 transition-colors cursor-pointer"
            aria-label="Close dialog"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* 2. Scrollable Body */}
        <div className="flex-1 overflow-y-auto p-5 sm:p-6 space-y-6">
          {/* Section A: Menu Identity & Branding */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 pb-4 border-b border-neutral-200/80">
            <div className="sm:col-span-2 space-y-1.5">
              <label htmlFor="input-menu-title" className="block text-xs font-semibold text-neutral-700">
                Menu Item Display Name <span className="text-red-500">*</span>
              </label>
              <input
                id="input-menu-title"
                type="text"
                value={displayName}
                onChange={(e) => {
                  setDisplayName(e.target.value);
                  setFormErrors((prev) => ({ ...prev, displayName: undefined }));
                }}
                placeholder="e.g., Executive 360: Projects & IT Fleet"
                className={`w-full px-3 py-2 text-sm bg-white border rounded-xl outline-none transition-all ${
                  formErrors.displayName
                    ? 'border-red-400 ring-2 ring-red-100'
                    : 'border-neutral-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-100'
                }`}
              />
              {formErrors.displayName && (
                <p className="text-xs text-red-500">{formErrors.displayName}</p>
              )}
            </div>

            {/* Icon & Color Accent */}
            <div className="space-y-1.5">
              <label className="block text-xs font-semibold text-neutral-700">Icon & Color Accent</label>
              <div className="flex items-center gap-2">
                {/* Icon Selector */}
                <select
                  value={selectedIcon}
                  onChange={(e) => setSelectedIcon(e.target.value)}
                  className="px-2.5 py-2 text-xs bg-white border border-neutral-300 rounded-xl focus:border-blue-500 focus:ring-2 focus:ring-blue-100 outline-none flex-1"
                >
                  {Object.entries(AVAILABLE_ICONS).map(([iconName, iconData]) => (
                    <option key={iconName} value={iconName}>
                      {iconData.label}
                    </option>
                  ))}
                </select>

                {/* Color Selector */}
                <div className="flex items-center gap-1 bg-neutral-100 p-1 rounded-xl border border-neutral-200">
                  {(['blue', 'indigo', 'emerald', 'purple', 'amber'] as ConnectionColor[]).map((c) => (
                    <button
                      key={c}
                      type="button"
                      onClick={() => setSelectedColor(c)}
                      className={`w-5 h-5 rounded-full transition-all cursor-pointer ${
                        COLOR_VARIANTS[c].badge
                      } ${selectedColor === c ? 'ring-2 ring-neutral-800 scale-110' : 'opacity-70 hover:opacity-100'}`}
                      title={c}
                    />
                  ))}
                </div>
              </div>
            </div>
          </div>

          {/* Section B: Connected SharePoint Sources Header & Actions */}
          <div className="space-y-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div>
                <div className="flex items-center gap-2">
                  <h3 className="text-sm font-semibold text-neutral-900">
                    Connected SharePoint Lists
                  </h3>
                  <span className="px-2 py-0.5 rounded-full text-xs font-bold bg-blue-100 text-blue-800 border border-blue-200">
                    {sources.length} {sources.length === 1 ? 'SharePoint List' : 'SharePoint Lists'} Connected
                  </span>
                  <span className="text-xs text-neutral-500 font-mono">
                    ({totalSelectedColumnsAcrossSources} total fields selected)
                  </span>
                </div>
                <p className="text-xs text-neutral-500 mt-0.5">
                  Connect multiple SharePoint sites/lists. Select the specific fields from each list to show side-by-side in this menu's table.
                </p>
              </div>

              {/* Connect Another SharePoint List button */}
              <button
                type="button"
                onClick={handleAddSource}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white text-xs font-semibold shadow-xs transition-all cursor-pointer flex-shrink-0"
              >
                <Plus className="w-4 h-4" />
                <span>Connect Another SharePoint List</span>
              </button>
            </div>

            {formErrors.sources && (
              <p className="text-xs text-red-500 bg-red-50 p-2 rounded-lg border border-red-200">
                {formErrors.sources}
              </p>
            )}

            {/* Sources Tabs / Pills for Fast Switching & Visibility */}
            <div className="flex items-center gap-2 overflow-x-auto pb-1 pt-1 border-b border-neutral-200">
              {sources.map((src, index) => {
                const isActive = index === activeSourceIndex;
                const fieldCount = src.selectedColumnNames?.length || 0;
                const preset = SHAREPOINT_SITE_PRESETS.find((s) => s.id === src.siteId);
                const tagColor = preset?.badgeColor || 'blue';

                return (
                  <button
                    key={src.id}
                    type="button"
                    onClick={() => setActiveSourceIndex(index)}
                    className={`inline-flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium border transition-all cursor-pointer flex-shrink-0 ${
                      isActive
                        ? 'bg-blue-50 text-blue-900 border-blue-300 shadow-xs ring-1 ring-blue-300'
                        : 'bg-neutral-50 text-neutral-600 border-neutral-200 hover:bg-neutral-100 hover:border-neutral-300'
                    }`}
                  >
                    <span
                      className={`w-2 h-2 rounded-full ${
                        tagColor === 'emerald'
                          ? 'bg-emerald-500'
                          : tagColor === 'purple'
                          ? 'bg-purple-500'
                          : tagColor === 'amber'
                          ? 'bg-amber-500'
                          : tagColor === 'cyan'
                          ? 'bg-cyan-500'
                          : 'bg-blue-500'
                      }`}
                    />
                    <span className="font-semibold">
                      SharePoint #{index + 1}:
                    </span>
                    <span className="truncate max-w-[130px]">
                      {src.listName}
                    </span>
                    <span className="px-1.5 py-0.2 rounded-full text-[10px] font-mono bg-white border border-neutral-200 text-neutral-600">
                      {fieldCount} fields
                    </span>

                    {/* Remove button if more than 1 source */}
                    {sources.length > 1 && (
                      <span
                        onClick={(e) => {
                          e.stopPropagation();
                          handleRemoveSource(index);
                        }}
                        className="ml-1 p-0.5 rounded-full hover:bg-neutral-200/80 text-neutral-400 hover:text-red-600"
                        title="Disconnect this SharePoint list"
                      >
                        <X className="w-3.5 h-3.5" />
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Section C: Active SharePoint Source Configuration Panel */}
          {currentActiveSource && (
            <div className="bg-neutral-50/70 rounded-2xl border border-neutral-200/90 p-4 sm:p-5 space-y-4">
              <div className="flex items-center justify-between gap-2 border-b border-neutral-200/70 pb-3">
                <div className="flex items-center gap-2">
                  <span className="px-2 py-0.5 rounded-md text-xs font-bold bg-neutral-900 text-white">
                    SharePoint Source #{activeSourceIndex + 1}
                  </span>
                  <span className="text-xs font-medium text-neutral-600">
                    Target Site & List
                  </span>
                </div>

                {sources.length > 1 && (
                  <button
                    type="button"
                    onClick={() => handleRemoveSource(activeSourceIndex)}
                    className="inline-flex items-center gap-1 text-xs text-red-600 hover:text-red-700 hover:bg-red-50 px-2 py-1 rounded-lg transition-colors cursor-pointer"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Disconnect this Source</span>
                  </button>
                )}
              </div>

              {/* Site & List Selectors for this Source */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {/* Target Site */}
                <div className="space-y-1.5">
                  <label className="block text-xs font-semibold text-neutral-700">
                    SharePoint Site Target
                  </label>
                  <select
                    value={currentActiveSource.siteId}
                    onChange={(e) => handleSourceSiteChange(activeSourceIndex, e.target.value)}
                    className="w-full px-3 py-2 text-xs bg-white border border-neutral-300 rounded-xl focus:border-blue-500 focus:ring-2 focus:ring-blue-100 outline-none"
                  >
                    {SHAREPOINT_SITE_PRESETS.map((site) => (
                      <option key={site.id} value={site.id}>
                        {site.name} ({site.shortCode})
                      </option>
                    ))}
                  </select>
                  <p className="text-[11px] text-neutral-400 font-mono truncate">
                    {currentActiveSource.siteUrl}
                  </p>
                </div>

                {/* Target List */}
                <div className="space-y-1.5">
                  <label className="block text-xs font-semibold text-neutral-700">
                    Target SharePoint List
                  </label>
                  <select
                    value={currentActiveSource.listId}
                    onChange={(e) => handleSourceListChange(activeSourceIndex, e.target.value)}
                    className="w-full px-3 py-2 text-xs bg-white border border-neutral-300 rounded-xl focus:border-blue-500 focus:ring-2 focus:ring-blue-100 outline-none"
                  >
                    {DEMO_SHAREPOINT_LISTS.map((list) => (
                      <option key={list.id} value={list.id}>
                        {list.displayName} ({list.columns.length} schema fields)
                      </option>
                    ))}
                  </select>
                  <p className="text-[11px] text-neutral-500 truncate">
                    {currentActiveListDef.description || 'Target SharePoint List'}
                  </p>
                </div>
              </div>

              {/* Field / Column Selection Toolbar for this Source */}
              <div className="pt-2 border-t border-neutral-200/70 space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                  <div>
                    <h4 className="text-xs font-bold uppercase tracking-wider text-neutral-700">
                      Select Fields from {currentActiveSource.listName}
                    </h4>
                    <p className="text-[11px] text-neutral-500">
                      {currentActiveSource.selectedColumnNames?.length || 0} of {currentActiveListDef.columns.length} fields selected from this list
                    </p>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => handleSelectAllColumns(activeSourceIndex)}
                      className="px-2 py-1 rounded-lg text-xs font-medium bg-white hover:bg-neutral-100 border border-neutral-200 text-neutral-700 cursor-pointer"
                    >
                      Select All
                    </button>
                    <button
                      type="button"
                      onClick={() => handleSelectQuickColumns(activeSourceIndex)}
                      className="px-2 py-1 rounded-lg text-xs font-medium bg-white hover:bg-neutral-100 border border-neutral-200 text-neutral-700 cursor-pointer"
                    >
                      First 5
                    </button>
                    <button
                      type="button"
                      onClick={() => handleDeselectAllColumns(activeSourceIndex)}
                      className="px-2 py-1 rounded-lg text-xs font-medium bg-white hover:bg-neutral-100 border border-neutral-200 text-neutral-500 hover:text-neutral-700 cursor-pointer"
                    >
                      Clear
                    </button>
                  </div>
                </div>

                {/* Field Search Filter */}
                <div className="relative">
                  <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
                  <input
                    type="text"
                    value={columnSearchQueries[currentActiveSource.id] || ''}
                    onChange={(e) =>
                      setColumnSearchQueries({
                        ...columnSearchQueries,
                        [currentActiveSource.id]: e.target.value,
                      })
                    }
                    placeholder={`Search available fields in ${currentActiveSource.listName}...`}
                    className="w-full pl-8 pr-3 py-1.5 text-xs bg-white border border-neutral-200 rounded-xl focus:border-blue-500 focus:ring-2 focus:ring-blue-100 outline-none"
                  />
                </div>

                {/* Available Fields Grid for this Source */}
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2 max-h-56 overflow-y-auto p-1 bg-white rounded-xl border border-neutral-200/80">
                  {filteredActiveColumns.map((colDef) => {
                    const isSelected = (currentActiveSource.selectedColumnNames || []).includes(
                      colDef.name
                    );
                    const fieldType = detectSharePointFieldType(colDef);

                    return (
                      <div
                        key={colDef.name}
                        onClick={() => handleToggleColumn(activeSourceIndex, colDef.name)}
                        className={`flex items-center justify-between p-2 rounded-lg border text-xs cursor-pointer select-none transition-all ${
                          isSelected
                            ? 'bg-blue-50/90 border-blue-300 text-blue-900 shadow-2xs font-medium'
                            : 'bg-white border-neutral-200/80 text-neutral-700 hover:bg-neutral-50'
                        }`}
                      >
                        <div className="flex items-center gap-2 min-w-0 pr-1">
                          <div
                            className={`w-4 h-4 rounded flex items-center justify-center border transition-all flex-shrink-0 ${
                              isSelected
                                ? 'bg-blue-600 border-blue-600 text-white'
                                : 'border-neutral-300 bg-white'
                            }`}
                          >
                            {isSelected && <Check className="w-3 h-3 stroke-[3]" />}
                          </div>
                          <span className="truncate" title={colDef.displayName}>
                            {colDef.displayName}
                          </span>
                        </div>

                        <span className="text-[10px] font-mono px-1 py-0.2 rounded bg-neutral-100 text-neutral-500 flex-shrink-0">
                          {fieldType}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          )}

          {/* Section D: Single Consolidated Table Columns Live Preview */}
          <div className="bg-neutral-100/70 rounded-2xl border border-neutral-200/80 p-4 space-y-2.5">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Layers className="w-4 h-4 text-blue-600" />
                <h4 className="text-xs font-bold uppercase tracking-wider text-neutral-800">
                  Single Table Combined Schema Preview
                </h4>
              </div>
              <span className="text-xs font-mono font-semibold text-blue-700 bg-blue-50 px-2 py-0.5 rounded-full border border-blue-200">
                {totalSelectedColumnsAcrossSources} total columns in single table
              </span>
            </div>

            <p className="text-[11px] text-neutral-500">
              The single table will display all selected fields from your {sources.length} connected SharePoint lists. The first column is locked sticky left and the last column is locked sticky right.
            </p>

            {/* Chips of combined columns */}
            <div className="flex flex-wrap gap-1.5 max-h-36 overflow-y-auto p-2 bg-white rounded-xl border border-neutral-200">
              {sources.flatMap((src, srcIdx) => {
                const preset = SHAREPOINT_SITE_PRESETS.find((s) => s.id === src.siteId);
                const tagColor = preset?.badgeColor || 'blue';
                const sList =
                  DEMO_SHAREPOINT_LISTS.find((l) => l.id === src.listId) || DEMO_SHAREPOINT_LISTS[0];

                return (src.selectedColumnNames || []).map((colName, cIdx) => {
                  const colDef = sList.columns.find((c) => c.name === colName);
                  const isOverallFirst = srcIdx === 0 && cIdx === 0;
                  const isOverallLast =
                    srcIdx === sources.length - 1 &&
                    cIdx === (src.selectedColumnNames?.length || 0) - 1 &&
                    totalSelectedColumnsAcrossSources > 1;

                  return (
                    <div
                      key={`${src.id}-${colName}`}
                      className="inline-flex items-center gap-1.5 px-2 py-1 rounded-lg text-xs bg-neutral-50 border border-neutral-200"
                    >
                      <span
                        className={`text-[9px] font-bold px-1.5 py-0.2 rounded-full uppercase ${
                          tagColor === 'emerald'
                            ? 'bg-emerald-100 text-emerald-800'
                            : tagColor === 'purple'
                            ? 'bg-purple-100 text-purple-800'
                            : tagColor === 'amber'
                            ? 'bg-amber-100 text-amber-800'
                            : tagColor === 'cyan'
                            ? 'bg-cyan-100 text-cyan-800'
                            : 'bg-blue-100 text-blue-800'
                        }`}
                      >
                        {src.listName.slice(0, 12)}
                      </span>
                      <span className="font-medium text-neutral-800">
                        {colDef?.displayName || colName}
                      </span>

                      {isOverallFirst && (
                        <span className="text-[9px] font-bold px-1 py-0.2 rounded bg-amber-100 text-amber-800 border border-amber-300">
                          Sticky Left
                        </span>
                      )}
                      {isOverallLast && (
                        <span className="text-[9px] font-bold px-1 py-0.2 rounded bg-indigo-100 text-indigo-800 border border-indigo-300">
                          Sticky Right
                        </span>
                      )}
                    </div>
                  );
                });
              })}
            </div>
          </div>
        </div>

        {/* 3. Modal Footer: Actions & Delete */}
        <div className="flex items-center justify-between px-6 py-4 border-t border-neutral-200 bg-neutral-50/80 flex-shrink-0">
          <div>
            {isEditing && (
              <>
                {!showDeleteConfirm ? (
                  <button
                    type="button"
                    onClick={() => setShowDeleteConfirm(true)}
                    className="text-xs font-medium text-red-600 hover:text-red-700 hover:bg-red-50 px-2.5 py-1.5 rounded-lg transition-colors cursor-pointer"
                  >
                    Delete Menu
                  </button>
                ) : (
                  <div className="flex items-center gap-2">
                    <span className="text-xs text-red-600 font-medium">Delete this menu?</span>
                    <button
                      type="button"
                      onClick={handleDelete}
                      className="px-2.5 py-1 text-xs font-semibold bg-red-600 text-white rounded-lg hover:bg-red-700 cursor-pointer"
                    >
                      Confirm
                    </button>
                    <button
                      type="button"
                      onClick={() => setShowDeleteConfirm(false)}
                      className="px-2 py-1 text-xs font-medium text-neutral-600 hover:bg-neutral-200 rounded-lg cursor-pointer"
                    >
                      Cancel
                    </button>
                  </div>
                )}
              </>
            )}
          </div>

          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={closeModal}
              className="px-4 py-2 text-xs font-semibold text-neutral-700 hover:bg-neutral-200/80 rounded-xl transition-colors cursor-pointer"
            >
              Cancel
            </button>
            <button
              id="btn-save-connection"
              type="button"
              onClick={handleSave}
              className="px-5 py-2 text-xs font-semibold bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white rounded-xl shadow-xs transition-all cursor-pointer"
            >
              {isEditing ? 'Save Changes' : 'Create Menu & Show Table'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
