import React, { useState, useEffect } from 'react';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { ScrollArea, ScrollBar } from '@/components/ui/scroll-area';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { 
  Database, 
  RefreshCw, 
  Trash2, 
  Eye, 
  ChevronDown,
  ChevronUp,
  Download,
  AlertCircle
} from 'lucide-react';
import { 
  getAllTables, 
  getTableData, 
  clearTable, 
  getTableStats
} from '@/lib/etl/database';
import { DatabaseRecord } from '@/lib/etl/types';
import { toast } from 'sonner';

interface TableMeta {
  tableName: string;
  headers: string[];
  recordCount: number;
  lastUpdated: string;
  createdAt: string;
}

export function DatabaseMonitor() {
  const [tables, setTables] = useState<TableMeta[]>([]);
  const [stats, setStats] = useState({ totalTables: 0, totalRecords: 0 });
  const [expandedTable, setExpandedTable] = useState<string | null>(null);
  const [tableData, setTableData] = useState<DatabaseRecord[]>([]);
  const [isLoading, setIsLoading] = useState(false);

  const loadTables = async () => {
    setIsLoading(true);
    try {
      const tablesData = await getAllTables();
      setTables(tablesData);
      const statsData = await getTableStats();
      setStats(statsData);
    } catch (error) {
      console.error('Failed to load tables:', error);
      toast.error('Gagal memuat data tabel');
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadTables();
  }, []);

  const handleViewTable = async (tableName: string) => {
    if (expandedTable === tableName) {
      setExpandedTable(null);
      setTableData([]);
      return;
    }

    try {
      const data = await getTableData(tableName);
      setTableData(data);
      setExpandedTable(tableName);
    } catch (error) {
      console.error('Failed to load table data:', error);
      toast.error('Gagal memuat data tabel');
    }
  };

  const handleClearTable = async (tableName: string) => {
    if (!confirm(`Apakah Anda yakin ingin menghapus semua data di tabel "${tableName}"?`)) {
      return;
    }

    try {
      await clearTable(tableName);
      toast.success(`Tabel "${tableName}" berhasil dikosongkan`);
      await loadTables();
      if (expandedTable === tableName) {
        setExpandedTable(null);
        setTableData([]);
      }
    } catch (error) {
      console.error('Failed to clear table:', error);
      toast.error('Gagal mengosongkan tabel');
    }
  };

  const handleExportTable = async (tableName: string) => {
    try {
      const data = await getTableData(tableName);
      const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${tableName}_${new Date().toISOString().split('T')[0]}.json`;
      a.click();
      URL.revokeObjectURL(url);
      toast.success('Data berhasil diekspor');
    } catch (error) {
      console.error('Failed to export table:', error);
      toast.error('Gagal mengekspor data');
    }
  };

  const formatDate = (dateStr: string) => {
    return new Date(dateStr).toLocaleString('id-ID', {
      dateStyle: 'medium',
      timeStyle: 'short',
    });
  };

  const formatCellValue = (value: unknown): string => {
    if (value === null || value === undefined) return '-';
    if (typeof value === 'number') {
      return value.toLocaleString('id-ID', { maximumFractionDigits: 2 });
    }
    return String(value);
  };

  if (tables.length === 0 && !isLoading) {
    return (
      <Card>
        <CardContent className="flex flex-col items-center justify-center py-12 text-center">
          <Database className="w-12 h-12 text-muted-foreground mb-4" />
          <h3 className="font-semibold text-lg mb-2">Database Kosong</h3>
          <p className="text-muted-foreground text-sm max-w-md">
            Belum ada data yang disimpan. Unggah file Excel dan klik "Simpan ke Database" untuk memulai.
          </p>
        </CardContent>
      </Card>
    );
  }

  return (
    <div className="space-y-4">
      {/* Stats Header */}
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <Database className="w-5 h-5 text-primary" />
            <span className="font-semibold">{stats.totalTables} Tabel</span>
          </div>
          <div className="text-muted-foreground">
            {stats.totalRecords.toLocaleString()} total record
          </div>
        </div>
        <Button variant="outline" size="sm" onClick={loadTables} disabled={isLoading}>
          <RefreshCw className={`w-4 h-4 mr-2 ${isLoading ? 'animate-spin' : ''}`} />
          Refresh
        </Button>
      </div>

      {/* Tables List */}
      <div className="space-y-3">
        {tables.map((table) => (
          <Card key={table.tableName}>
            <CardHeader className="py-3">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <CardTitle className="text-base font-mono">{table.tableName}</CardTitle>
                  <Badge variant="secondary">{table.recordCount} records</Badge>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-muted-foreground">
                    Update: {formatDate(table.lastUpdated)}
                  </span>
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => handleViewTable(table.tableName)}
                  >
                    {expandedTable === table.tableName ? (
                      <ChevronUp className="w-4 h-4" />
                    ) : (
                      <Eye className="w-4 h-4" />
                    )}
                  </Button>
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => handleExportTable(table.tableName)}
                  >
                    <Download className="w-4 h-4" />
                  </Button>
                  <Button
                    variant="ghost"
                    size="sm"
                    className="text-destructive hover:text-destructive"
                    onClick={() => handleClearTable(table.tableName)}
                  >
                    <Trash2 className="w-4 h-4" />
                  </Button>
                </div>
              </div>
            </CardHeader>

            {expandedTable === table.tableName && tableData.length > 0 && (
              <CardContent className="pt-0">
                <ScrollArea className="h-[300px] rounded-lg border">
                  <div className="min-w-max">
                    <Table>
                      <TableHeader className="sticky top-0 bg-muted/95 backdrop-blur-sm">
                        <TableRow>
                          {Object.keys(tableData[0])
                            .filter((k) => k !== '_id')
                            .map((key) => (
                              <TableHead key={key} className="whitespace-nowrap min-w-[100px]">
                                {key}
                              </TableHead>
                            ))}
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {tableData.slice(0, 50).map((row, idx) => (
                          <TableRow key={idx}>
                            {Object.entries(row)
                              .filter(([k]) => k !== '_id')
                              .map(([key, value], colIdx) => (
                                <TableCell
                                  key={colIdx}
                                  className="whitespace-nowrap max-w-[200px] truncate"
                                >
                                  {formatCellValue(value)}
                                </TableCell>
                              ))}
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                  <ScrollBar orientation="horizontal" />
                </ScrollArea>
                {tableData.length > 50 && (
                  <p className="text-sm text-muted-foreground text-center mt-2">
                    Menampilkan 50 dari {tableData.length} records
                  </p>
                )}
              </CardContent>
            )}
          </Card>
        ))}
      </div>
    </div>
  );
}
