import React, { useState, useEffect } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { 
  Upload, 
  FileSpreadsheet, 
  CheckCircle, 
  AlertCircle,
  Database,
  RefreshCw,
  Loader2
} from 'lucide-react';
import { 
  loadMasterDataFromFile, 
  loadMasterDataFromPublic, 
  isMasterLoaded,
  getMasterDataStats,
  clearMasterData
} from '@/lib/etl/masterData';
import { toast } from 'sonner';

interface MasterDataUploadProps {
  onMasterDataLoaded?: () => void;
}

export function MasterDataUpload({ onMasterDataLoaded }: MasterDataUploadProps) {
  const [isLoading, setIsLoading] = useState(false);
  const [isLoaded, setIsLoaded] = useState(false);
  const [stats, setStats] = useState<{
    totalEmiten: number;
    uniqueGroups: number;
    categories: Record<string, number>;
  } | null>(null);
  const [loadSource, setLoadSource] = useState<'default' | 'custom' | null>(null);

  // Try to load default master data on mount
  useEffect(() => {
    if (!isMasterLoaded()) {
      loadDefaultMasterData();
    } else {
      setIsLoaded(true);
      setStats(getMasterDataStats());
      setLoadSource('default');
    }
  }, []);

  const loadDefaultMasterData = async () => {
    setIsLoading(true);
    try {
      const result = await loadMasterDataFromPublic();
      if (result.success) {
        setIsLoaded(true);
        setStats(getMasterDataStats());
        setLoadSource('default');
        onMasterDataLoaded?.();
      }
    } catch (error) {
      console.warn('Could not load default master data:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    try {
      const result = await loadMasterDataFromFile(file);
      
      if (result.success) {
        setIsLoaded(true);
        setStats(getMasterDataStats());
        setLoadSource('custom');
        toast.success(`Master data berhasil dimuat: ${result.count} emiten`);
        onMasterDataLoaded?.();
      } else {
        toast.error(result.errors.join(', '));
      }
    } catch (error) {
      toast.error('Gagal memuat master data');
    } finally {
      setIsLoading(false);
      // Reset input
      event.target.value = '';
    }
  };

  const handleReset = () => {
    clearMasterData();
    setIsLoaded(false);
    setStats(null);
    setLoadSource(null);
    loadDefaultMasterData();
  };

  return (
    <Card>
      <CardHeader className="pb-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Database className="w-5 h-5 text-primary" />
            <CardTitle className="text-lg">Master Data Emiten</CardTitle>
          </div>
          {isLoaded && (
            <Badge variant="outline" className="gap-1">
              <CheckCircle className="w-3 h-3" />
              {loadSource === 'default' ? 'Default' : 'Custom'}
            </Badge>
          )}
        </div>
        <CardDescription>
          Data referensi untuk pencocokan Kode Efek ke Grup Afiliasi
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        {isLoading ? (
          <div className="flex items-center justify-center py-4">
            <Loader2 className="w-6 h-6 animate-spin text-muted-foreground" />
            <span className="ml-2 text-sm text-muted-foreground">Memuat master data...</span>
          </div>
        ) : isLoaded && stats ? (
          <>
            <Alert>
              <CheckCircle className="w-4 h-4" />
              <AlertTitle>Master Data Aktif</AlertTitle>
              <AlertDescription className="mt-2">
                <div className="grid grid-cols-2 gap-4 text-sm">
                  <div>
                    <span className="text-muted-foreground">Total Emiten:</span>
                    <span className="ml-2 font-medium">{stats.totalEmiten.toLocaleString()}</span>
                  </div>
                  <div>
                    <span className="text-muted-foreground">Grup Unik:</span>
                    <span className="ml-2 font-medium">{stats.uniqueGroups}</span>
                  </div>
                </div>
                <div className="mt-3 flex flex-wrap gap-1">
                  {Object.entries(stats.categories).slice(0, 5).map(([cat, count]) => (
                    <Badge key={cat} variant="secondary" className="text-xs">
                      {cat}: {count}
                    </Badge>
                  ))}
                </div>
              </AlertDescription>
            </Alert>
            
            <div className="flex gap-2">
              <label className="flex-1">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileUpload}
                  className="hidden"
                />
                <Button variant="outline" className="w-full gap-2" asChild>
                  <span>
                    <Upload className="w-4 h-4" />
                    Ganti Master Data
                  </span>
                </Button>
              </label>
              <Button variant="ghost" size="icon" onClick={handleReset}>
                <RefreshCw className="w-4 h-4" />
              </Button>
            </div>
          </>
        ) : (
          <>
            <Alert variant="destructive">
              <AlertCircle className="w-4 h-4" />
              <AlertTitle>Master Data Tidak Tersedia</AlertTitle>
              <AlertDescription>
                Upload file master data emiten untuk mengaktifkan fitur grouping dan agregasi.
              </AlertDescription>
            </Alert>
            
            <label>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileUpload}
                className="hidden"
              />
              <Button className="w-full gap-2" asChild>
                <span>
                  <FileSpreadsheet className="w-4 h-4" />
                  Upload Master Data
                </span>
              </Button>
            </label>
          </>
        )}
      </CardContent>
    </Card>
  );
}
