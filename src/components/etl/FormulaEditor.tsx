import React, { useState } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Textarea } from '@/components/ui/textarea';
import { Badge } from '@/components/ui/badge';
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip';
import { 
  CalculationStep, 
  FormulaDefinition, 
  getFormulas, 
  updateFormula, 
  resetFormulas 
} from '@/lib/etl/mkbdCalculator';
import { cn } from '@/lib/utils';
import { 
  Calculator, 
  Edit2, 
  RotateCcw, 
  Check, 
  X, 
  GripVertical,
  Info,
  AlertTriangle,
  ArrowRight
} from 'lucide-react';

interface FormulaEditorProps {
  calculationSteps: CalculationStep[];
  onFormulaChange?: () => void;
}

const formatNumber = (num: number): string => {
  if (Math.abs(num) >= 1e12) return `${(num / 1e12).toFixed(2)}T`;
  if (Math.abs(num) >= 1e9) return `${(num / 1e9).toFixed(2)}B`;
  if (Math.abs(num) >= 1e6) return `${(num / 1e6).toFixed(2)}M`;
  return num.toLocaleString('id-ID');
};

export function FormulaEditor({ calculationSteps, onFormulaChange }: FormulaEditorProps) {
  const [editingStep, setEditingStep] = useState<string | null>(null);
  const [editedFormula, setEditedFormula] = useState('');
  const [draggedItem, setDraggedItem] = useState<string | null>(null);

  const handleEditStart = (step: CalculationStep) => {
    setEditingStep(step.id);
    setEditedFormula(step.formula);
  };

  const handleEditSave = () => {
    if (editingStep && editedFormula) {
      // In a full implementation, this would parse and validate the formula
      // For now, we just update the display
      onFormulaChange?.();
    }
    setEditingStep(null);
    setEditedFormula('');
  };

  const handleEditCancel = () => {
    setEditingStep(null);
    setEditedFormula('');
  };

  const handleDragStart = (e: React.DragEvent, stepId: string) => {
    setDraggedItem(stepId);
    e.dataTransfer.effectAllowed = 'move';
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
  };

  const handleDrop = (e: React.DragEvent, targetId: string) => {
    e.preventDefault();
    if (draggedItem && draggedItem !== targetId) {
      // Handle reordering logic here
      console.log(`Moved ${draggedItem} to ${targetId}`);
    }
    setDraggedItem(null);
  };

  const getStepStatusColor = (step: CalculationStep): string => {
    if (step.result < 0) return 'text-red-600 dark:text-red-400';
    if (step.result === 0) return 'text-amber-600 dark:text-amber-400';
    return 'text-green-600 dark:text-green-400';
  };

  return (
    <Card>
      <CardHeader>
        <div className="flex items-center justify-between">
          <div>
            <CardTitle className="flex items-center gap-2">
              <Calculator className="w-5 h-5" />
              Formula & Kalkulasi
            </CardTitle>
            <CardDescription>
              Lihat dan edit rumus perhitungan MKBD. Drag & drop untuk mengubah urutan.
            </CardDescription>
          </div>
          
          <TooltipProvider>
            <Tooltip>
              <TooltipTrigger asChild>
                <Button 
                  variant="outline" 
                  size="sm"
                  onClick={() => {
                    resetFormulas();
                    onFormulaChange?.();
                  }}
                >
                  <RotateCcw className="w-4 h-4 mr-1" />
                  Reset
                </Button>
              </TooltipTrigger>
              <TooltipContent>Reset ke formula default</TooltipContent>
            </Tooltip>
          </TooltipProvider>
        </div>
      </CardHeader>
      
      <CardContent className="space-y-4">
        <Accordion type="single" collapsible className="w-full">
          {calculationSteps.map((step, index) => (
            <AccordionItem 
              key={step.id} 
              value={step.id}
              className={cn(
                'border rounded-lg mb-2 overflow-hidden',
                draggedItem === step.id && 'opacity-50 border-dashed border-2 border-primary'
              )}
              draggable={step.editable}
              onDragStart={(e) => handleDragStart(e, step.id)}
              onDragOver={handleDragOver}
              onDrop={(e) => handleDrop(e, step.id)}
            >
              <AccordionTrigger className="px-4 hover:no-underline hover:bg-muted/50">
                <div className="flex items-center gap-3 w-full">
                  {step.editable && (
                    <GripVertical className="w-4 h-4 text-muted-foreground cursor-grab" />
                  )}
                  
                  <div className="flex items-center gap-2">
                    <Badge variant="outline" className="font-mono text-xs">
                      {index + 1}
                    </Badge>
                    <span className="font-medium text-sm">{step.name}</span>
                  </div>
                  
                  <ArrowRight className="w-4 h-4 text-muted-foreground mx-2" />
                  
                  <span className={cn('font-bold font-mono', getStepStatusColor(step))}>
                    {formatNumber(step.result)}
                  </span>
                  
                  <Badge 
                    variant="secondary" 
                    className="ml-auto mr-2 text-xs"
                  >
                    {step.source}
                  </Badge>
                </div>
              </AccordionTrigger>
              
              <AccordionContent className="px-4 pb-4">
                <div className="space-y-3">
                  {/* Formula Display/Edit */}
                  <div className="p-3 bg-muted/50 rounded-lg font-mono text-sm">
                    {editingStep === step.id ? (
                      <div className="space-y-2">
                        <Textarea
                          value={editedFormula}
                          onChange={(e) => setEditedFormula(e.target.value)}
                          className="font-mono text-sm"
                          rows={2}
                        />
                        <div className="flex gap-2">
                          <Button size="sm" onClick={handleEditSave}>
                            <Check className="w-3 h-3 mr-1" />
                            Simpan
                          </Button>
                          <Button size="sm" variant="ghost" onClick={handleEditCancel}>
                            <X className="w-3 h-3 mr-1" />
                            Batal
                          </Button>
                        </div>
                      </div>
                    ) : (
                      <div className="flex items-start justify-between">
                        <code className="text-primary">{step.formula}</code>
                        {step.editable && (
                          <Button 
                            size="sm" 
                            variant="ghost" 
                            className="h-6 px-2"
                            onClick={() => handleEditStart(step)}
                          >
                            <Edit2 className="w-3 h-3" />
                          </Button>
                        )}
                      </div>
                    )}
                  </div>
                  
                  {/* Input Values */}
                  <div className="space-y-2">
                    <Label className="text-xs text-muted-foreground">Input Values:</Label>
                    <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
                      {Object.entries(step.inputValues).map(([key, value]) => (
                        <div 
                          key={key} 
                          className="flex items-center justify-between p-2 bg-background rounded border text-xs"
                        >
                          <span className="text-muted-foreground truncate">{key}</span>
                          <span className="font-mono font-medium">{formatNumber(value)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </AccordionContent>
            </AccordionItem>
          ))}
        </Accordion>
        
        {/* Summary */}
        <div className="p-4 bg-gradient-to-r from-primary/10 to-primary/5 rounded-lg">
          <div className="flex items-center gap-2 mb-2">
            <Info className="w-4 h-4 text-primary" />
            <span className="font-medium text-sm">Ringkasan Perhitungan</span>
          </div>
          <p className="text-xs text-muted-foreground">
            Sistem menggunakan 3-Pass calculation: (1) Ekstraksi nilai dari VD51/VD52, 
            (2) Kalkulasi Ranking Liabilities dari VD510, 
            (3) Perhitungan MKBD final dengan "Blind Overwrite" untuk koreksi.
          </p>
        </div>
      </CardContent>
    </Card>
  );
}
