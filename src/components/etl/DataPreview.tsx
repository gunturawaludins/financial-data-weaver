import React from 'react';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { ScrollArea, ScrollBar } from '@/components/ui/scroll-area';
import { Badge } from '@/components/ui/badge';
import { ProcessedSheet } from '@/lib/etl/types';

interface DataPreviewProps {
  sheet: ProcessedSheet;
  maxRows?: number;
}

export function DataPreview({ sheet, maxRows = 100 }: DataPreviewProps) {
  const displayData = sheet.data.slice(0, maxRows);
  const displayHeaders = sheet.headers.filter(h => !h.startsWith('_'));
  
  const formatCellValue = (value: unknown): string => {
    if (value === null || value === undefined) return '-';
    if (typeof value === 'number') {
      // Format numbers with thousand separators
      return value.toLocaleString('id-ID', { maximumFractionDigits: 2 });
    }
    return String(value);
  };

  return (
    <div className="space-y-3">
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
          {sheet.metadata.originalRowCount !== sheet.metadata.cleanedRowCount && (
            <>
              <span>•</span>
              <span className="text-amber-600">
                {sheet.metadata.originalRowCount - sheet.metadata.cleanedRowCount} baris dihapus
              </span>
            </>
          )}
        </div>
      </div>

      <ScrollArea className="h-[400px] rounded-lg border">
        <div className="min-w-max">
          <Table>
            <TableHeader className="sticky top-0 bg-muted/95 backdrop-blur-sm z-10">
              <TableRow>
                <TableHead className="w-12 text-center font-bold">#</TableHead>
                {displayHeaders.map((header, idx) => (
                  <TableHead 
                    key={idx} 
                    className="whitespace-nowrap font-semibold min-w-[120px]"
                  >
                    {header}
                  </TableHead>
                ))}
              </TableRow>
            </TableHeader>
            <TableBody>
              {displayData.map((row, rowIdx) => (
                <TableRow key={rowIdx} className="hover:bg-muted/50">
                  <TableCell className="text-center text-muted-foreground font-mono text-xs">
                    {rowIdx + 1}
                  </TableCell>
                  {displayHeaders.map((header, colIdx) => (
                    <TableCell 
                      key={colIdx} 
                      className="whitespace-nowrap max-w-[300px] truncate"
                      title={formatCellValue(row[header])}
                    >
                      {formatCellValue(row[header])}
                    </TableCell>
                  ))}
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
        <ScrollBar orientation="horizontal" />
      </ScrollArea>

      {sheet.data.length > maxRows && (
        <p className="text-sm text-muted-foreground text-center">
          Menampilkan {maxRows} dari {sheet.data.length} baris
        </p>
      )}
    </div>
  );
}
