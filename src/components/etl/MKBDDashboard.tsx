import React from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Progress } from '@/components/ui/progress';
import { Separator } from '@/components/ui/separator';
import { MKBDCalculationResult } from '@/lib/etl/mkbdCalculator';
import { cn } from '@/lib/utils';
import { 
  TrendingUp, 
  TrendingDown, 
  AlertTriangle, 
  CheckCircle2,
  Banknote,
  Scale,
  ShieldCheck,
  ArrowUpRight,
  ArrowDownRight
} from 'lucide-react';

interface MKBDDashboardProps {
  result: MKBDCalculationResult;
}

const formatNumber = (num: number): string => {
  if (Math.abs(num) >= 1e12) return `Rp ${(num / 1e12).toFixed(2)} T`;
  if (Math.abs(num) >= 1e9) return `Rp ${(num / 1e9).toFixed(2)} B`;
  if (Math.abs(num) >= 1e6) return `Rp ${(num / 1e6).toFixed(2)} M`;
  return `Rp ${num.toLocaleString('id-ID')}`;
};

const formatShortNumber = (num: number): string => {
  if (Math.abs(num) >= 1e12) return `${(num / 1e12).toFixed(1)}T`;
  if (Math.abs(num) >= 1e9) return `${(num / 1e9).toFixed(1)}B`;
  if (Math.abs(num) >= 1e6) return `${(num / 1e6).toFixed(1)}M`;
  return num.toLocaleString('id-ID');
};

export function MKBDDashboard({ result }: MKBDDashboardProps) {
  const isHealthy = result.lebihKurangMKBD >= 0;
  const mkbdRatio = (result.mkbdDisesuaikan / result.mkbdDiwajibkan) * 100;
  const rankingLiabilitiesRatio = result.totalRankingLiabilities > 0 
    ? (result.totalRankingLiabilities / (result.totalEkuitas || 1)) * 100 
    : 0;

  return (
    <div className="space-y-6">
      {/* Hero Card - MKBD Status */}
      <Card className={cn(
        'border-2',
        isHealthy ? 'border-green-500/30 bg-gradient-to-br from-green-50 to-emerald-50 dark:from-green-950/20 dark:to-emerald-950/20' 
                  : 'border-red-500/30 bg-gradient-to-br from-red-50 to-rose-50 dark:from-red-950/20 dark:to-rose-950/20'
      )}>
        <CardHeader>
          <div className="flex items-center justify-between">
            <div>
              <CardTitle className="flex items-center gap-2">
                {isHealthy ? (
                  <CheckCircle2 className="w-6 h-6 text-green-600" />
                ) : (
                  <AlertTriangle className="w-6 h-6 text-red-600" />
                )}
                Status MKBD
              </CardTitle>
              <CardDescription>
                Modal Kerja Bersih Disesuaikan
              </CardDescription>
            </div>
            <Badge 
              variant={isHealthy ? 'default' : 'destructive'}
              className="text-lg px-4 py-1"
            >
              {isHealthy ? 'SEHAT' : 'KRITIS'}
            </Badge>
          </div>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {/* MKBD Disesuaikan */}
            <div className="space-y-2">
              <p className="text-sm text-muted-foreground">MKBD Disesuaikan</p>
              <p className="text-3xl font-bold text-foreground">
                {formatShortNumber(result.mkbdDisesuaikan)}
              </p>
              <div className="flex items-center gap-1 text-sm">
                {result.mkbdDisesuaikan > result.mkbdDiwajibkan ? (
                  <>
                    <ArrowUpRight className="w-4 h-4 text-green-600" />
                    <span className="text-green-600">Di atas minimum</span>
                  </>
                ) : (
                  <>
                    <ArrowDownRight className="w-4 h-4 text-red-600" />
                    <span className="text-red-600">Di bawah minimum</span>
                  </>
                )}
              </div>
            </div>

            {/* MKBD Diwajibkan */}
            <div className="space-y-2">
              <p className="text-sm text-muted-foreground">MKBD Diwajibkan</p>
              <p className="text-3xl font-bold text-muted-foreground">
                {formatShortNumber(result.mkbdDiwajibkan)}
              </p>
              <p className="text-sm text-muted-foreground">
                6.25% × (Liabilitas + Ranking)
              </p>
            </div>

            {/* Lebih/Kurang */}
            <div className="space-y-2">
              <p className="text-sm text-muted-foreground">Lebih/(Kurang)</p>
              <p className={cn(
                'text-3xl font-bold',
                isHealthy ? 'text-green-600' : 'text-red-600'
              )}>
                {result.lebihKurangMKBD >= 0 ? '+' : ''}{formatShortNumber(result.lebihKurangMKBD)}
              </p>
              <Progress 
                value={Math.min(mkbdRatio, 200)} 
                className={cn(
                  'h-2',
                  isHealthy ? '[&>div]:bg-green-500' : '[&>div]:bg-red-500'
                )}
              />
              <p className="text-xs text-muted-foreground text-right">
                {mkbdRatio.toFixed(1)}% dari minimum
              </p>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Key Metrics Grid */}
      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        {/* Total Aset Lancar */}
        <Card>
          <CardContent className="pt-6">
            <div className="flex items-center justify-between mb-2">
              <Banknote className="w-5 h-5 text-blue-500" />
              <Badge variant="outline" className="text-xs">VD59</Badge>
            </div>
            <p className="text-xs text-muted-foreground">Total Aset Lancar</p>
            <p className="text-xl font-bold">{formatShortNumber(result.totalAsetLancar)}</p>
          </CardContent>
        </Card>

        {/* Total Ekuitas */}
        <Card>
          <CardContent className="pt-6">
            <div className="flex items-center justify-between mb-2">
              <ShieldCheck className="w-5 h-5 text-green-500" />
              <Badge variant="outline" className="text-xs">VD52</Badge>
            </div>
            <p className="text-xs text-muted-foreground">Total Ekuitas</p>
            <p className="text-xl font-bold">{formatShortNumber(result.totalEkuitas)}</p>
          </CardContent>
        </Card>

        {/* Total Liabilitas */}
        <Card>
          <CardContent className="pt-6">
            <div className="flex items-center justify-between mb-2">
              <Scale className="w-5 h-5 text-amber-500" />
              <Badge variant="outline" className="text-xs">VD59</Badge>
            </div>
            <p className="text-xs text-muted-foreground">Total Liabilitas</p>
            <p className="text-xl font-bold">{formatShortNumber(result.totalLiabilitas)}</p>
          </CardContent>
        </Card>

        {/* Total Ranking Liabilities */}
        <Card>
          <CardContent className="pt-6">
            <div className="flex items-center justify-between mb-2">
              <TrendingDown className="w-5 h-5 text-red-500" />
              <Badge variant="outline" className="text-xs">VD510</Badge>
            </div>
            <p className="text-xs text-muted-foreground">Ranking Liabilities</p>
            <p className="text-xl font-bold text-red-600">
              {formatShortNumber(result.totalRankingLiabilities)}
            </p>
            <p className="text-xs text-muted-foreground">
              {rankingLiabilitiesRatio.toFixed(1)}% dari aset
            </p>
          </CardContent>
        </Card>

        {/* Modal Kerja */}
        <Card>
          <CardContent className="pt-6">
            <div className="flex items-center justify-between mb-2">
              <ShieldCheck className="w-5 h-5 text-green-500" />
              <Badge variant="outline" className="text-xs">Calc</Badge>
            </div>
            <p className="text-xs text-muted-foreground">Modal Kerja</p>
            <p className="text-xl font-bold">{formatShortNumber(result.modalKerja)}</p>
          </CardContent>
        </Card>
      </div>

      {/* VD510 Details */}
      {result.vd510Details.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle className="text-lg">Detail Ranking Liabilities (VD510)</CardTitle>
            <CardDescription>
              Perhitungan per emiten yang melebihi 20% dari Total Ekuitas
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b">
                    <th className="text-left py-2 px-2 font-medium text-muted-foreground">#</th>
                    <th className="text-left py-2 px-2 font-medium text-muted-foreground">Kode</th>
                    <th className="text-left py-2 px-2 font-medium text-muted-foreground">Grup Emiten</th>
                    <th className="text-right py-2 px-2 font-medium text-muted-foreground">Nilai Pasar</th>
                    <th className="text-right py-2 px-2 font-medium text-muted-foreground">% Modal</th>
                    <th className="text-right py-2 px-2 font-medium text-muted-foreground">Ranking Liab.</th>
                  </tr>
                </thead>
                <tbody>
                  {result.vd510Details
                    .filter(d => d.nilaiRankingLiabilities > 0)
                    .slice(0, 10)
                    .map((detail, idx) => (
                      <tr key={idx} className="border-b hover:bg-muted/50">
                        <td className="py-2 px-2 text-muted-foreground">{idx + 1}</td>
                        <td className="py-2 px-2 font-mono font-medium">{detail.kodeEfek}</td>
                        <td className="py-2 px-2">{detail.grupEmiten}</td>
                        <td className="py-2 px-2 text-right font-mono">
                          {formatShortNumber(detail.nilaiPasarWajar)}
                        </td>
                        <td className="py-2 px-2 text-right">
                          <Badge variant={detail.persentaseTerhadapModal > 20 ? 'destructive' : 'secondary'}>
                            {detail.persentaseTerhadapModal.toFixed(1)}%
                          </Badge>
                        </td>
                        <td className="py-2 px-2 text-right font-mono text-red-600">
                          {formatShortNumber(detail.nilaiRankingLiabilities)}
                        </td>
                      </tr>
                    ))}
                </tbody>
                <tfoot>
                  <tr className="bg-muted/50 font-bold">
                    <td colSpan={5} className="py-2 px-2 text-right">Total Ranking Liabilities:</td>
                    <td className="py-2 px-2 text-right font-mono text-red-600">
                      {formatShortNumber(result.totalRankingLiabilities)}
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </CardContent>
        </Card>
      )}

      {/* VD59 Updates */}
      {result.vd59Updates && result.vd59Updates.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle className="text-lg">Update VD59 (Calculated Values)</CardTitle>
            <CardDescription>
              Nilai yang diperbarui berdasarkan perhitungan sistem
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b">
                    <th className="text-left py-2 px-2 font-medium text-muted-foreground">Baris</th>
                    <th className="text-left py-2 px-2 font-medium text-muted-foreground">Keterangan</th>
                    <th className="text-left py-2 px-2 font-medium text-muted-foreground">Kolom</th>
                    <th className="text-right py-2 px-2 font-medium text-muted-foreground">Nilai Lama</th>
                    <th className="text-right py-2 px-2 font-medium text-muted-foreground">Nilai Baru</th>
                    <th className="text-left py-2 px-2 font-medium text-muted-foreground">Formula</th>
                  </tr>
                </thead>
                <tbody>
                  {result.vd59Updates.map((update, idx) => (
                    <tr key={idx} className="border-b hover:bg-muted/50">
                      <td className="py-2 px-2 font-mono">{update.rowIndex}</td>
                      <td className="py-2 px-2 text-xs max-w-[200px] truncate" title={update.rowDescription}>
                        {update.rowDescription}
                      </td>
                      <td className="py-2 px-2 font-mono text-xs">{update.column}</td>
                      <td className="py-2 px-2 text-right font-mono text-muted-foreground">
                        {formatShortNumber(update.oldValue)}
                      </td>
                      <td className="py-2 px-2 text-right font-mono font-bold text-primary">
                        {formatShortNumber(update.newValue)}
                      </td>
                      <td className="py-2 px-2 text-xs text-muted-foreground max-w-[200px] truncate" title={update.formula}>
                        {update.formula}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Calculation Flow */}
      <Card>
        <CardHeader>
          <CardTitle className="text-lg">Alur Perhitungan</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="space-y-4">
            <div className="flex flex-wrap items-center gap-2 text-sm">
              <Badge variant="outline" className="font-mono">
                Aset Lancar: {formatShortNumber(result.totalAsetLancar)}
              </Badge>
              <span className="text-muted-foreground">−</span>
              <Badge variant="outline" className="font-mono">
                Liabilitas: {formatShortNumber(result.totalLiabilitas)}
              </Badge>
              <span className="text-muted-foreground">−</span>
              <Badge variant="destructive" className="font-mono">
                Ranking: {formatShortNumber(result.totalRankingLiabilities)}
              </Badge>
              <span className="text-muted-foreground">=</span>
              <Badge variant="default" className="font-mono">
                Modal Kerja: {formatShortNumber(result.modalKerja)}
              </Badge>
            </div>
            
            {result.haircutSum > 0 && (
              <div className="flex flex-wrap items-center gap-2 text-sm">
                <Badge variant="default" className="font-mono">
                  Modal Kerja: {formatShortNumber(result.modalKerja)}
                </Badge>
                <span className="text-muted-foreground">−</span>
                <Badge variant="secondary" className="font-mono">
                  Haircut (33-92): {formatShortNumber(result.haircutSum)}
                </Badge>
                <span className="text-muted-foreground">=</span>
                <Badge variant="default" className="font-mono bg-green-600">
                  MKBD Disesuaikan: {formatShortNumber(result.mkbdDisesuaikan)}
                </Badge>
              </div>
            )}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
