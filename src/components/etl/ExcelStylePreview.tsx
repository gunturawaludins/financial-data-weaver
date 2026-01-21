import React, { useState, useMemo } from 'react';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { ScrollArea, ScrollBar } from '@/components/ui/scroll-area';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip';
import { ProcessedSheet } from '@/lib/etl/types';
import { cn } from '@/lib/utils';
import { Search, ChevronLeft, ChevronRight, Grid3X3, Eye } from 'lucide-react';

interface ExcelStylePreviewProps {
  sheet: ProcessedSheet;
  maxRows?: number;
  showEmptyBlocks?: boolean;
  highlightFormulas?: boolean;
}

// Excel column letters
const getExcelColumn = (index: number): string => {
  let column = '';
  let temp = index;
  while (temp >= 0) {
    column = String.fromCharCode((temp % 26) + 65) + column;
    temp = Math.floor(temp / 26) - 1;
  }
  return column;
};

export function ExcelStylePreview({ 
  sheet, 
  maxRows = 100,
  showEmptyBlocks = true,
  highlightFormulas = true,
}: ExcelStylePreviewProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(0);
  const [showGridLines, setShowGridLines] = useState(true);
  const rowsPerPage = 50;

  const displayHeaders = useMemo(() => 
    sheet.headers.filter(h => !h.startsWith('_')), 
    [sheet.headers]
  );
  
  const filteredData = useMemo(() => {
    if (!searchTerm) return sheet.data;
    
    const lowerSearch = searchTerm.toLowerCase();
    return sheet.data.filter(row => 
      Object.values(row).some(val => 
        String(val ?? '').toLowerCase().includes(lowerSearch)
      )
    );
  }, [sheet.data, searchTerm]);

  const paginatedData = useMemo(() => {
    const start = currentPage * rowsPerPage;
    return filteredData.slice(start, start + rowsPerPage);
  }, [filteredData, currentPage, rowsPerPage]);

  const totalPages = Math.ceil(filteredData.length / rowsPerPage);

  const formatCellValue = (value: unknown): string => {
    if (value === null || value === undefined || value === '') return '';
    if (typeof value === 'number') {
      // Format numbers with thousand separators
      if (Math.abs(value) >= 1e12) {
        return `${(value / 1e12).toFixed(2)}T`;
      }
      if (Math.abs(value) >= 1e9) {
        return `${(value / 1e9).toFixed(2)}B`;
      }
      if (Math.abs(value) >= 1e6) {
        return `${(value / 1e6).toFixed(2)}M`;
      }
      return value.toLocaleString('id-ID', { maximumFractionDigits: 2 });
    }
    return String(value);
  };

  const getCellStyle = (value: unknown, header: string): string => {
    const baseStyle = 'px-2 py-1.5 text-xs font-mono';
    const isEmpty = value === null || value === undefined || value === '';
    
    if (isEmpty && showEmptyBlocks) {
      return cn(baseStyle, 'bg-muted/30');
    }
    
    if (typeof value === 'number') {
      return cn(baseStyle, 'text-right', value < 0 ? 'text-red-600' : 'text-foreground');
    }
    
    // Highlight formula-like values
    if (highlightFormulas && typeof value === 'string') {
      if (value.includes('=') || value.includes('SUM') || value.includes('MAX')) {
        return cn(baseStyle, 'bg-blue-50 text-blue-700 dark:bg-blue-900/20 dark:text-blue-400');
      }
    }
    
    // Highlight important rows
    const strValue = String(value).toLowerCase();
    if (strValue.includes('total') || strValue.includes('jumlah')) {
      return cn(baseStyle, 'font-bold bg-amber-50 dark:bg-amber-900/20');
    }
    
    return baseStyle;
  };

  const isRowMerged = (row: Record<string, unknown>): boolean => {
    // Check if most cells are empty (merged row indicator)
    const nonEmptyCount = Object.values(row).filter(v => 
      v !== null && v !== undefined && String(v).trim() !== ''
    ).length;
    return nonEmptyCount <= 2;
  };

  return (
    <div className="space-y-3">
      {/* Header Bar */}
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div className="flex items-center gap-2">
          <h3 className="font-semibold text-foreground">{sheet.sheetName}</h3>
          <Badge variant="secondary" className="font-mono text-xs">
            {sheet.tableName}
          </Badge>
        </div>
        
        <div className="flex items-center gap-3 text-sm text-muted-foreground">
          <span>{sheet.metadata.cleanedRowCount} baris</span>
          <span>•</span>
          <span>{displayHeaders.length} kolom</span>
        </div>
      </div>

      {/* Toolbar */}
      <div className="flex items-center gap-2 flex-wrap">
        <div className="relative flex-1 min-w-[200px] max-w-sm">
          <Search className="absolute left-2 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
          <Input
            placeholder="Cari data..."
            value={searchTerm}
            onChange={(e) => {
              setSearchTerm(e.target.value);
              setCurrentPage(0);
            }}
            className="pl-8 h-8 text-sm"
          />
        </div>
        
        <TooltipProvider>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant={showGridLines ? 'secondary' : 'ghost'}
                size="sm"
                onClick={() => setShowGridLines(!showGridLines)}
              >
                <Grid3X3 className="w-4 h-4" />
              </Button>
            </TooltipTrigger>
            <TooltipContent>Toggle grid lines</TooltipContent>
          </Tooltip>
        </TooltipProvider>

        <TooltipProvider>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant={highlightFormulas ? 'secondary' : 'ghost'}
                size="sm"
              >
                <Eye className="w-4 h-4" />
              </Button>
            </TooltipTrigger>
            <TooltipContent>Show formula highlights</TooltipContent>
          </Tooltip>
        </TooltipProvider>
      </div>

      {/* Excel-style Table */}
      <ScrollArea className="h-[500px] rounded-lg border">
        <div className="min-w-max">
          <Table className={cn(!showGridLines && '[&_td]:border-0 [&_th]:border-0')}>
            {/* Column Headers (A, B, C...) */}
            <TableHeader className="sticky top-0 z-20">
              <TableRow className="bg-muted/80 backdrop-blur-sm">
                <TableHead className="w-12 text-center font-bold text-muted-foreground bg-muted/95 border-r sticky left-0 z-30">
                  
                </TableHead>
                {displayHeaders.map((_, idx) => (
                  <TableHead 
                    key={idx} 
                    className="text-center text-xs font-bold text-muted-foreground bg-muted/95 min-w-[80px] border-b"
                  >
                    {getExcelColumn(idx)}
                  </TableHead>
                ))}
              </TableRow>
              {/* Actual Headers */}
              <TableRow className="bg-slate-100 dark:bg-slate-800">
                <TableHead className="w-12 text-center font-bold text-muted-foreground bg-slate-100 dark:bg-slate-800 border-r sticky left-0 z-30">
                  #
                </TableHead>
                {displayHeaders.map((header, idx) => (
                  <TableHead 
                    key={idx} 
                    className="whitespace-nowrap text-xs font-semibold min-w-[100px] max-w-[200px] truncate px-2"
                    title={header}
                  >
                    {header}
                  </TableHead>
                ))}
              </TableRow>
            </TableHeader>
            
            <TableBody>
              {paginatedData.map((row, rowIdx) => {
                const actualRowNum = currentPage * rowsPerPage + rowIdx + 1;
                const isMergedRow = isRowMerged(row);
                
                return (
                  <TableRow 
                    key={rowIdx} 
                    className={cn(
                      'hover:bg-blue-50/50 dark:hover:bg-blue-900/20',
                      isMergedRow && 'bg-slate-50 dark:bg-slate-800/50'
                    )}
                  >
                    {/* Row Number */}
                    <TableCell className="text-center text-muted-foreground font-mono text-xs bg-muted/30 border-r sticky left-0 z-10">
                      {actualRowNum}
                    </TableCell>
                    
                    {/* Data Cells */}
                    {displayHeaders.map((header, colIdx) => {
                      const value = row[header];
                      const isEmpty = value === null || value === undefined || value === '';
                      
                      return (
                        <TableCell 
                          key={colIdx} 
                          className={cn(
                            getCellStyle(value, header),
                            isEmpty && showEmptyBlocks && 'bg-gradient-to-r from-muted/40 to-muted/20',
                            'max-w-[200px] truncate'
                          )}
                          title={isEmpty ? '(kosong)' : formatCellValue(value)}
                        >
                          {isEmpty && showEmptyBlocks ? (
                            <span className="text-muted-foreground/30">—</span>
                          ) : (
                            formatCellValue(value)
                          )}
                        </TableCell>
                      );
                    })}
                  </TableRow>
                );
              })}
            </TableBody>
          </Table>
        </div>
        <ScrollBar orientation="horizontal" />
      </ScrollArea>

      {/* Pagination */}
      {totalPages > 1 && (
        <div className="flex items-center justify-between">
          <p className="text-sm text-muted-foreground">
            Menampilkan {currentPage * rowsPerPage + 1} - {Math.min((currentPage + 1) * rowsPerPage, filteredData.length)} dari {filteredData.length} baris
          </p>
          
          <div className="flex items-center gap-2">
            <Button
              variant="outline"
              size="sm"
              disabled={currentPage === 0}
              onClick={() => setCurrentPage(p => p - 1)}
            >
              <ChevronLeft className="w-4 h-4" />
            </Button>
            <span className="text-sm text-muted-foreground">
              {currentPage + 1} / {totalPages}
            </span>
            <Button
              variant="outline"
              size="sm"
              disabled={currentPage >= totalPages - 1}
              onClick={() => setCurrentPage(p => p + 1)}
            >
              <ChevronRight className="w-4 h-4" />
            </Button>
          </div>
        </div>
      )}
    </div>
  );
}
