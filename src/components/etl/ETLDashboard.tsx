import React, { useState } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Badge } from '@/components/ui/badge';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { 
  Upload, 
  Database, 
  CheckCircle, 
  AlertCircle, 
  FileSpreadsheet,
  Save,
  Loader2,
  Info,
  Users,
  TrendingUp
} from 'lucide-react';
import { FileUpload } from './FileUpload';
import { DataPreview } from './DataPreview';
import { DatabaseMonitor } from './DatabaseMonitor';
import { MasterDataUpload } from './MasterDataUpload';
import { extractFromExcel, ProcessedSheet, ETLResult } from '@/lib/etl';
import { createTableIfNotExists, appendRecords } from '@/lib/etl/database';
import { DatabaseRecord } from '@/lib/etl/types';
import { toast } from 'sonner';

export function ETLDashboard() {
  const [activeTab, setActiveTab] = useState('upload');
  const [isProcessing, setIsProcessing] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [etlResult, setEtlResult] = useState<ETLResult | null>(null);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState(0);
  const [fileName, setFileName] = useState<string>('');
  const [masterDataVersion, setMasterDataVersion] = useState(0);

  const handleFileSelect = async (file: File) => {
    setIsProcessing(true);
    setEtlResult(null);
    setFileName(file.name);

    try {
      const result = await extractFromExcel(file);
      setEtlResult(result);
      setSelectedSheetIndex(0);
      
      if (result.success && result.sheets.length > 0) {
        toast.success(`${result.sheets.length} sheet berhasil diproses`);
      } else if (!result.success) {
        toast.error('Gagal memproses file');
      }
    } catch (error) {
      console.error('ETL Error:', error);
      toast.error('Terjadi kesalahan saat memproses file');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleSaveToDatabase = async () => {
    if (!etlResult?.sheets.length) return;

    setIsSaving(true);
    let savedCount = 0;
    let totalRecords = 0;

    try {
      for (const sheet of etlResult.sheets) {
        // Create table if not exists
        await createTableIfNotExists(sheet.tableName, sheet.headers);
        
        // Append records
        const count = await appendRecords(sheet.tableName, sheet.data as DatabaseRecord[]);
        savedCount++;
        totalRecords += count;
      }

      toast.success(
        `Berhasil menyimpan ${totalRecords.toLocaleString()} record ke ${savedCount} tabel`
      );
      
      // Switch to monitor tab
      setActiveTab('monitor');
    } catch (error) {
      console.error('Save Error:', error);
      toast.error('Gagal menyimpan data ke database');
    } finally {
      setIsSaving(false);
    }
  };

  const currentSheet = etlResult?.sheets[selectedSheetIndex];

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="border-b bg-card">
        <div className="container mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="p-2 rounded-lg bg-primary/10">
                <FileSpreadsheet className="w-6 h-6 text-primary" />
              </div>
              <div>
                <h1 className="text-xl font-bold text-foreground">
                  ETL Laporan MKBD
                </h1>
                <p className="text-sm text-muted-foreground">
                  Extract, Transform, Load untuk Laporan Keuangan
                </p>
              </div>
            </div>
            {fileName && (
              <Badge variant="outline" className="font-mono">
                {fileName}
              </Badge>
            )}
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="container mx-auto px-4 py-6">
        <Tabs value={activeTab} onValueChange={setActiveTab} className="space-y-6">
          <TabsList className="grid w-full max-w-md grid-cols-2">
            <TabsTrigger value="upload" className="flex items-center gap-2">
              <Upload className="w-4 h-4" />
              Upload & Preview
            </TabsTrigger>
            <TabsTrigger value="monitor" className="flex items-center gap-2">
              <Database className="w-4 h-4" />
              Database Monitor
            </TabsTrigger>
          </TabsList>

          {/* Upload Tab */}
          <TabsContent value="upload" className="space-y-6">
            {/* Master Data Section */}
            <MasterDataUpload 
              onMasterDataLoaded={() => setMasterDataVersion(v => v + 1)} 
            />

            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <Upload className="w-5 h-5" />
                  Upload File Laporan MKBD
                </CardTitle>
                <CardDescription>
                  Unggah file Excel laporan MKBD. Sistem akan otomatis:
                  <ul className="list-disc list-inside mt-2 space-y-1">
                    <li>Memindai semua sheet dan membersihkan data</li>
                    <li>Mencocokkan Kode Efek dengan Master Data (VLOOKUP)</li>
                    <li>Menghitung GRUP NILAI PASAR WAJAR (Agregasi)</li>
                  </ul>
                </CardDescription>
              </CardHeader>
              <CardContent>
                <FileUpload onFileSelect={handleFileSelect} isLoading={isProcessing} />
              </CardContent>
            </Card>

            {/* Processing Info */}
            {etlResult && (
              <>
                {/* Status Alert */}
                <Alert variant={etlResult.success ? 'default' : 'destructive'}>
                  {etlResult.success ? (
                    <CheckCircle className="w-4 h-4" />
                  ) : (
                    <AlertCircle className="w-4 h-4" />
                  )}
                  <AlertTitle>
                    {etlResult.success ? 'Berhasil Diproses' : 'Proses Gagal'}
                  </AlertTitle>
                  <AlertDescription>
                    {etlResult.success
                      ? `${etlResult.sheets.length} sheet siap untuk disimpan ke database.`
                      : etlResult.errors.join(', ')}
                  </AlertDescription>
                </Alert>

                {/* Warnings */}
                {etlResult.warnings.length > 0 && (
                  <Alert>
                    <Info className="w-4 h-4" />
                    <AlertTitle>Informasi</AlertTitle>
                    <AlertDescription>
                      <ul className="list-disc list-inside space-y-1 mt-2">
                        {etlResult.warnings.slice(0, 5).map((warning, idx) => (
                          <li key={idx} className="text-sm">{warning}</li>
                        ))}
                        {etlResult.warnings.length > 5 && (
                          <li className="text-sm text-muted-foreground">
                            ...dan {etlResult.warnings.length - 5} peringatan lainnya
                          </li>
                        )}
                      </ul>
                    </AlertDescription>
                  </Alert>
                )}

                {/* Enrichment Stats */}
                {currentSheet?.metadata.enrichmentStats && (
                  <Alert>
                    <TrendingUp className="w-4 h-4" />
                    <AlertTitle>Statistik Enrichment & Agregasi</AlertTitle>
                    <AlertDescription>
                      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-2">
                        <div className="text-center p-2 bg-muted rounded">
                          <div className="text-lg font-bold text-primary">
                            {currentSheet.metadata.enrichmentStats.matchedCount}
                          </div>
                          <div className="text-xs text-muted-foreground">Matched</div>
                        </div>
                        <div className="text-center p-2 bg-muted rounded">
                          <div className="text-lg font-bold text-orange-500">
                            {currentSheet.metadata.enrichmentStats.unmatchedCount}
                          </div>
                          <div className="text-xs text-muted-foreground">Non-Grup</div>
                        </div>
                        <div className="text-center p-2 bg-muted rounded">
                          <div className="text-lg font-bold text-blue-500">
                            {currentSheet.metadata.enrichmentStats.groupCount}
                          </div>
                          <div className="text-xs text-muted-foreground">Grup Unik</div>
                        </div>
                        <div className="text-center p-2 bg-muted rounded">
                          <div className="text-lg font-bold text-green-500">
                            {(currentSheet.metadata.enrichmentStats.totalGroupValue / 1e9).toFixed(2)}B
                          </div>
                          <div className="text-xs text-muted-foreground">Total Nilai</div>
                        </div>
                      </div>
                      {currentSheet.metadata.enrichmentStats.kodeEfekColumn && (
                        <p className="text-xs text-muted-foreground mt-2">
                          Kolom Kode: <Badge variant="outline">{currentSheet.metadata.enrichmentStats.kodeEfekColumn}</Badge>
                          {currentSheet.metadata.enrichmentStats.nilaiPasarColumn && (
                            <> • Kolom Nilai: <Badge variant="outline">{currentSheet.metadata.enrichmentStats.nilaiPasarColumn}</Badge></>
                          )}
                        </p>
                      )}
                    </AlertDescription>
                  </Alert>
                )}

                {/* Sheet Tabs */}
                {etlResult.sheets.length > 0 && (
                  <Card>
                    <CardHeader className="pb-3">
                      <div className="flex items-center justify-between flex-wrap gap-4">
                        <CardTitle>Preview Data Bersih</CardTitle>
                        <Button 
                          onClick={handleSaveToDatabase}
                          disabled={isSaving}
                          className="gap-2"
                        >
                          {isSaving ? (
                            <Loader2 className="w-4 h-4 animate-spin" />
                          ) : (
                            <Save className="w-4 h-4" />
                          )}
                          Simpan ke Database
                        </Button>
                      </div>
                    </CardHeader>
                    <CardContent className="space-y-4">
                      {/* Sheet Selector */}
                      {etlResult.sheets.length > 1 && (
                        <div className="flex flex-wrap gap-2">
                          {etlResult.sheets.map((sheet, idx) => (
                            <Button
                              key={idx}
                              variant={selectedSheetIndex === idx ? 'default' : 'outline'}
                              size="sm"
                              onClick={() => setSelectedSheetIndex(idx)}
                            >
                              {sheet.sheetName}
                              <Badge 
                                variant={selectedSheetIndex === idx ? 'secondary' : 'outline'}
                                className="ml-2"
                              >
                                {sheet.data.length}
                              </Badge>
                            </Button>
                          ))}
                        </div>
                      )}

                      {/* Data Preview */}
                      {currentSheet && <DataPreview sheet={currentSheet} />}
                    </CardContent>
                  </Card>
                )}
              </>
            )}
          </TabsContent>

          {/* Database Monitor Tab */}
          <TabsContent value="monitor">
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <Database className="w-5 h-5" />
                  Database Monitor
                </CardTitle>
                <CardDescription>
                  Lihat dan kelola data yang tersimpan di database lokal (IndexedDB).
                  Data bersifat append (ditambahkan), tidak menimpa data lama.
                </CardDescription>
              </CardHeader>
              <CardContent>
                <DatabaseMonitor />
              </CardContent>
            </Card>
          </TabsContent>
        </Tabs>
      </main>

      {/* Footer */}
      <footer className="border-t mt-8">
        <div className="container mx-auto px-4 py-4">
          <p className="text-center text-sm text-muted-foreground">
            Sistem ETL Laporan MKBD • Data disimpan secara lokal di browser Anda
          </p>
        </div>
      </footer>
    </div>
  );
}
