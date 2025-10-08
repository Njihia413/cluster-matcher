// app/page.tsx
'use client';

import React, { useState } from 'react';
import { AlertCircle, CheckCircle, Upload, Download, FileText, Settings, X } from 'lucide-react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Badge } from '@/components/ui/badge';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';

interface ClusterFile {
  name: string;
  data: Record<string, unknown>[];
  headers: string[];
  columnsToAdd: string[];
  clusterMatchField: string;
  topologyMatchField: string;
}

interface ProcessedCluster {
  name: string;
  data: Record<string, unknown>[];
  columnSummaries: Record<string, Record<string, number>>;
  addedColumns: string[];
  clusterMatchField: string;
  topologyMatchField: string;
  headers: string[];
}

export default function ClusterMatcher() {
  const [topologyData, setTopologyData] = useState<Record<string, unknown>[]>([]);
  const [topologyColumns, setTopologyColumns] = useState<string[]>([]);
  const [clusterFiles, setClusterFiles] = useState<ClusterFile[]>([]);
  const [processedClusters, setProcessedClusters] = useState<ProcessedCluster[]>([]);
  const [globalMatchField, setGlobalMatchField] = useState('');
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');

  const handleTopologyFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setStatus('Loading topology tracker...');
      setError('');
      
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
      
      if (jsonData.length > 0) {
        const columns = Object.keys(jsonData[0] as Record<string, unknown>);
        setTopologyColumns(columns);
        setTopologyData(jsonData as Record<string, unknown>[]);
        setStatus(`✓ Loaded ${jsonData.length} records from topology tracker`);
        toast.success(`Topology tracker loaded: ${jsonData.length} records`);
      }
    } catch (err) {
      const error = err as Error;
      setError(`Error loading topology file: ${error.message}`);
      toast.error('Failed to load topology tracker');
    }
  };

  const handleClusterFiles = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    try {
      setStatus(`Loading ${files.length} cluster file(s)...`);
      setError('');
      
      const loadedFiles: ClusterFile[] = [];
      
      for (const file of files) {
        const text = await file.text();
        
        await new Promise<void>((resolve) => {
          Papa.parse(text, {
            header: true,
            dynamicTyping: true,
            skipEmptyLines: true,
            complete: (results) => {
              loadedFiles.push({
                name: file.name,
                data: results.data as Record<string, unknown>[],
                headers: results.meta.fields || [],
                columnsToAdd: [],
                clusterMatchField: '',
                topologyMatchField: ''
              });
              resolve();
            }
          });
        });
      }
      
      setClusterFiles(loadedFiles);
      setStatus(`✓ Loaded ${loadedFiles.length} cluster file(s)`);
      toast.success(`${loadedFiles.length} cluster file(s) loaded`);
    } catch (err) {
      const error = err as Error;
      setError(`Error loading cluster files: ${error.message}`);
      toast.error('Failed to load cluster files');
    }
  };

  const updateClusterConfig = (index: number, field: keyof ClusterFile, value: string) => {
    const updated = [...clusterFiles];
    (updated[index][field] as string) = value;
    setClusterFiles(updated);
  };

  const addColumnToCluster = (index: number, column: string) => {
    const updated = [...clusterFiles];
    if (!updated[index].columnsToAdd.includes(column)) {
      updated[index].columnsToAdd.push(column);
      setClusterFiles(updated);
    }
  };

  const removeColumnFromCluster = (clusterIndex: number, column: string) => {
    const updated = [...clusterFiles];
    updated[clusterIndex].columnsToAdd = updated[clusterIndex].columnsToAdd.filter(c => c !== column);
    setClusterFiles(updated);
  };

  const applyGlobalMatchField = () => {
    if (!globalMatchField) return;
    const updated = clusterFiles.map(cf => ({
      ...cf,
      clusterMatchField: globalMatchField,
      topologyMatchField: globalMatchField
    }));
    setClusterFiles(updated);
    setStatus(`✓ Applied "${globalMatchField}" as match field to all clusters`);
    toast.success(`Applied "${globalMatchField}" to all clusters`);
  };

  const processAllClusters = () => {
    // Validate configurations
    for (let i = 0; i < clusterFiles.length; i++) {
      const cf = clusterFiles[i];
      if (!cf.clusterMatchField || !cf.topologyMatchField) {
        setError(`Please select both match fields for "${cf.name}"`);
        toast.error(`Missing match fields for "${cf.name}"`);
        return;
      }
      if (cf.columnsToAdd.length === 0) {
        setError(`Please add at least one column for "${cf.name}"`);
        toast.error(`No columns selected for "${cf.name}"`);
        return;
      }
    }

    if (topologyData.length === 0) {
      setError('Please upload topology tracker file');
      toast.error('Topology tracker not loaded');
      return;
    }

    setStatus('Processing clusters...');
    
    const processed: ProcessedCluster[] = clusterFiles.map(cluster => {
      // Find the actual column name in topology data (case-insensitive)
      const topologyMatchField = topologyColumns.find(
        col => col.toLowerCase() === cluster.topologyMatchField.toLowerCase()
      ) || cluster.topologyMatchField;
      
      // Create lookup map for this cluster's match field
      const lookupMap: Record<string, any> = {};
      topologyData.forEach(row => {
        const key = (row[topologyMatchField] || '').toString().trim().toLowerCase();
        if (key) {
          lookupMap[key] = row;
        }
      });
      
      // Process data and add selected columns
      const matchedData = cluster.data.map(row => {
        const matchValue = (row[cluster.clusterMatchField] || '').toString().trim().toLowerCase();
        const match = lookupMap[matchValue];
        
        const newRow = { ...row };
        
        // Add each selected column
        cluster.columnsToAdd.forEach(col => {
          // Find actual column name in topology (case-insensitive)
          const actualCol = topologyColumns.find(
            c => c.toLowerCase() === col.toLowerCase()
          ) || col;
          
          const lookupValue = match 
            ? (match[actualCol] !== undefined && match[actualCol] !== '' ? match[actualCol] : 'Not Found')
            : 'Not Found';
          newRow[col] = lookupValue;
        });
        
        return newRow;
      });
      
      // Create summaries for each added column
      const columnSummaries: Record<string, Record<string, number>> = {};
      cluster.columnsToAdd.forEach(col => {
        const counts: Record<string, number> = {};
        matchedData.forEach(row => {
          const val = String(row[col] ?? 'Unknown');
          counts[val] = (counts[val] || 0) + 1;
        });
        columnSummaries[col] = counts;
      });
      
      return {
        name: cluster.name,
        data: matchedData,
        columnSummaries: columnSummaries,
        addedColumns: cluster.columnsToAdd,
        clusterMatchField: cluster.clusterMatchField,
        topologyMatchField: cluster.topologyMatchField,
        headers: [...cluster.headers, ...cluster.columnsToAdd.filter(c => !cluster.headers.includes(c))]
      };
    });
    
    setProcessedClusters(processed);
    setStatus(`✓ Processed ${processed.length} cluster file(s) successfully!`);
    toast.success(`Successfully processed ${processed.length} cluster file(s)`);
  };

  const downloadCluster = (cluster: ProcessedCluster) => {
    try {
      const csv = Papa.unparse(cluster.data);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const baseName = cluster.name.replace('.csv', '');
      const columnsStr = cluster.addedColumns.join('_');
      a.download = `${baseName}_with_${columnsStr}.csv`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
      
      toast.success(`Downloaded: ${a.download}`);
    } catch {
      toast.error(`Failed to download ${cluster.name}`);
    }
  };

  const downloadAll = () => {
    let successCount = 0;
    processedClusters.forEach((cluster, idx) => {
      setTimeout(() => {
        try {
          downloadCluster(cluster);
          successCount++;
          if (idx === processedClusters.length - 1) {
            toast.success(`All ${successCount} files downloaded successfully!`);
          }
        } catch {
          toast.error(`Failed to download ${cluster.name}`);
        }
      }, idx * 300);
    });
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 dark:from-slate-900 dark:to-slate-800 p-6">
      <div className="max-w-7xl mx-auto">
        <Card className="mb-6">
          <CardHeader>
            <CardTitle className="text-3xl">Cluster Region Matcher</CardTitle>
            <CardDescription>
              Match and enrich cluster files with data from topology tracker
            </CardDescription>
          </CardHeader>
        </Card>

        <div className="grid md:grid-cols-2 gap-6 mb-6">
          <Card>
            <CardHeader>
              <CardTitle>Step 1: Upload Topology Tracker</CardTitle>
            </CardHeader>
            <CardContent>
              <label className="cursor-pointer">
                <Button className="w-full" asChild>
                  <div className="flex items-center justify-center gap-2">
                    <Upload size={20} />
                    <span>Choose Excel File</span>
                  </div>
                </Button>
                <input 
                  type="file" 
                  accept=".xlsx,.xls" 
                  onChange={handleTopologyFile}
                  className="hidden"
                />
              </label>
              {topologyData.length > 0 && (
                <div className="mt-4 flex items-center gap-2 text-green-600">
                  <CheckCircle size={20} />
                  <span className="text-sm">Loaded {topologyData.length} records</span>
                </div>
              )}
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Step 2: Upload Cluster Files</CardTitle>
            </CardHeader>
            <CardContent>
              <label className="cursor-pointer">
                <Button variant="secondary" className="w-full" asChild>
                  <div className="flex items-center justify-center gap-2">
                    <FileText size={20} />
                    <span>Choose CSV File(s)</span>
                  </div>
                </Button>
                <input 
                  type="file" 
                  accept=".csv" 
                  multiple
                  onChange={handleClusterFiles}
                  className="hidden"
                />
              </label>
              {clusterFiles.length > 0 && (
                <div className="mt-4 flex items-center gap-2 text-green-600">
                  <CheckCircle size={20} />
                  <span className="text-sm">{clusterFiles.length} file(s) loaded</span>
                </div>
              )}
            </CardContent>
          </Card>
        </div>

        {clusterFiles.length > 0 && topologyColumns.length > 0 && (
          <>
            <Card className="mb-6">
              <CardHeader>
                <CardTitle>Quick Setup (Optional)</CardTitle>
                <CardDescription>
                  Apply the same match field to all cluster files
                </CardDescription>
              </CardHeader>
              <CardContent>
                <div className="flex gap-4">
                  <Select value={globalMatchField} onValueChange={setGlobalMatchField}>
                    <SelectTrigger className="flex-1">
                      <SelectValue placeholder="Select match field..." />
                    </SelectTrigger>
                    <SelectContent>
                      {clusterFiles[0]?.headers.map(col => (
                        <SelectItem key={col} value={col}>{col}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                  <Button 
                    onClick={applyGlobalMatchField}
                    disabled={!globalMatchField}
                  >
                    Apply to All
                  </Button>
                </div>
              </CardContent>
            </Card>

            <Card className="mb-6">
              <CardHeader>
                <div className="flex items-center gap-2">
                  <Settings size={24} />
                  <CardTitle>Step 3: Configure Each Cluster File</CardTitle>
                </div>
              </CardHeader>
              <CardContent>
                <div className="space-y-6">
                  {clusterFiles.map((cluster, idx) => (
                    <Card key={idx} className="bg-slate-50 dark:bg-slate-900">
                      <CardHeader>
                        <CardTitle className="text-lg">{cluster.name}</CardTitle>
                      </CardHeader>
                      <CardContent className="space-y-4">
                        <Alert>
                          <AlertDescription>
                            <p className="font-medium mb-3">Field Mapping</p>
                            <div className="grid md:grid-cols-2 gap-4">
                              <div>
                                <label className="block text-xs font-medium mb-2">
                                  Cluster File Field
                                </label>
                                <Select
                                  value={cluster.clusterMatchField}
                                  onValueChange={(value) => updateClusterConfig(idx, 'clusterMatchField', value)}
                                >
                                  <SelectTrigger>
                                    <SelectValue placeholder="Select from cluster..." />
                                  </SelectTrigger>
                                  <SelectContent>
                                    {cluster.headers.map(col => (
                                      <SelectItem key={col} value={col}>{col}</SelectItem>
                                    ))}
                                  </SelectContent>
                                </Select>
                              </div>

                              <div>
                                <label className="block text-xs font-medium mb-2">
                                  Topology Tracker Field
                                </label>
                                <Select
                                  value={cluster.topologyMatchField}
                                  onValueChange={(value) => updateClusterConfig(idx, 'topologyMatchField', value)}
                                >
                                  <SelectTrigger>
                                    <SelectValue placeholder="Select from topology..." />
                                  </SelectTrigger>
                                  <SelectContent>
                                    {topologyColumns.map(col => (
                                      <SelectItem key={col} value={col}>{col}</SelectItem>
                                    ))}
                                  </SelectContent>
                                </Select>
                              </div>
                            </div>
                            <p className="text-xs text-muted-foreground mt-2">
                              Example: Match cluster&apos;s &quot;nodeName&quot; with topology&apos;s &quot;SITENAME&quot;
                            </p>
                          </AlertDescription>
                        </Alert>

                        <div>
                          <label className="block text-sm font-medium mb-2">
                            Add Column from Topology
                          </label>
                          <Select
                            onValueChange={(value) => {
                              if (value) {
                                addColumnToCluster(idx, value);
                              }
                            }}
                          >
                            <SelectTrigger>
                              <SelectValue placeholder="Select column to add..." />
                            </SelectTrigger>
                            <SelectContent>
                              {topologyColumns.map(col => (
                                <SelectItem key={col} value={col}>{col}</SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                        </div>

                        {cluster.columnsToAdd.length > 0 && (
                          <div>
                            <label className="block text-sm font-medium mb-2">
                              Columns to Add:
                            </label>
                            <div className="flex flex-wrap gap-2">
                              {cluster.columnsToAdd.map(col => (
                                <Badge key={col} variant="secondary" className="gap-2">
                                  {col}
                                  <button
                                    onClick={() => removeColumnFromCluster(idx, col)}
                                    className="hover:bg-slate-200 dark:hover:bg-slate-700 rounded-full"
                                  >
                                    <X size={14} />
                                  </button>
                                </Badge>
                              ))}
                            </div>
                          </div>
                        )}
                      </CardContent>
                    </Card>
                  ))}
                </div>

                <Button 
                  onClick={processAllClusters}
                  className="w-full mt-6"
                  size="lg"
                >
                  <Settings className="mr-2" size={20} />
                  Process All Clusters
                </Button>
              </CardContent>
            </Card>
          </>
        )}

        {status && (
          <Alert className="mb-6">
            <AlertCircle className="h-4 w-4" />
            <AlertDescription>{status}</AlertDescription>
          </Alert>
        )}

        {error && (
          <Alert variant="destructive" className="mb-6">
            <AlertCircle className="h-4 w-4" />
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}

        {processedClusters.length > 0 && (
          <>
            <Card className="mb-6">
              <CardHeader>
                <div className="flex justify-between items-center">
                  <CardTitle>Download Results</CardTitle>
                  <Button onClick={downloadAll}>
                    <Download className="mr-2" size={20} />
                    Download All
                  </Button>
                </div>
              </CardHeader>
              <CardContent>
                <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-4">
                  {processedClusters.map((cluster, idx) => (
                    <Card key={idx}>
                      <CardHeader>
                        <CardTitle className="text-sm truncate">{cluster.name}</CardTitle>
                        <CardDescription className="text-xs">
                          {cluster.data.length} sites • Added: {cluster.addedColumns.join(', ')}
                        </CardDescription>
                      </CardHeader>
                      <CardContent>
                        <Button 
                          onClick={() => downloadCluster(cluster)}
                          className="w-full"
                          size="sm"
                        >
                          <Download className="mr-2" size={16} />
                          Download
                        </Button>
                      </CardContent>
                    </Card>
                  ))}
                </div>
              </CardContent>
            </Card>

            {processedClusters.map((cluster, idx) => (
              <Card key={idx} className="mb-6">
                <CardHeader>
                  <CardTitle>{cluster.name}</CardTitle>
                </CardHeader>
                <CardContent>
                  {cluster.addedColumns.map(col => (
                    <div key={col} className="mb-6">
                      <h3 className="text-lg font-medium mb-3">{col} Summary</h3>
                      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                        {Object.entries(cluster.columnSummaries[col] || {}).map(([value, count]) => (
                          <div 
                            key={value} 
                            className={`rounded-lg p-3 border ${
                              value === 'Not Found' 
                                ? 'bg-red-50 border-red-200 dark:bg-red-900/20' 
                                : 'bg-blue-50 border-blue-200 dark:bg-blue-900/20'
                            }`}
                          >
                            <div className={`text-xl font-bold ${
                              value === 'Not Found' ? 'text-red-600' : 'text-blue-600'
                            }`}>
                              {count}
                            </div>
                            <div className="text-sm font-medium truncate" title={value}>
                              {value}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}

                  <div className="overflow-x-auto">
                    <Table>
                      <TableHeader>
                        <TableRow>
                          {cluster.headers.map(header => (
                            <TableHead key={header} className={cluster.addedColumns.includes(header) ? 'font-bold' : ''}>
                              {header} {cluster.addedColumns.includes(header) && '✨'}
                            </TableHead>
                          ))}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {cluster.data.slice(0, 10).map((row, rowIdx) => (
                          <TableRow key={rowIdx}>
                            {cluster.headers.map(header => (
                              <TableCell key={header}>
                                <span className={
                                  cluster.addedColumns.includes(header) && row[header] === 'Not Found'
                                    ? 'text-red-600 font-medium'
                                    : cluster.addedColumns.includes(header)
                                    ? 'text-blue-600 font-medium'
                                    : ''
                                }>
                                  {row[header] !== undefined && row[header] !== null ? row[header].toString() : '-'}
                                </span>
                              </TableCell>
                            ))}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                    {cluster.data.length > 10 && (
                      <p className="text-muted-foreground text-center py-3 text-sm">
                        Showing 10 of {cluster.data.length} sites. Download CSV to see all.
                      </p>
                    )}
                  </div>
                </CardContent>
              </Card>
            ))}
          </>
        )}
      </div>
    </div>
  );
}
