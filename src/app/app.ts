import { Component, computed, signal } from '@angular/core';

import { ConverterError } from './converter/converter-error';
import { ConverterService, ConversionResult } from './converter/converter.service';

type Status = 'idle' | 'parsing' | 'done' | 'error';

const PREVIEW_LINES = 50;

@Component({
  selector: 'app-root',
  imports: [],
  templateUrl: './app.html',
  styleUrl: './app.css',
})
export class App {
  protected readonly status = signal<Status>('idle');
  protected readonly result = signal<ConversionResult | null>(null);
  protected readonly errorMessage = signal('');
  protected readonly dragging = signal(false);
  protected readonly previewOpen = signal(false);
  protected readonly copied = signal(false);

  protected readonly preview = computed(() => {
    const xml = this.result()?.xml;
    if (!xml) return '';
    const lines = xml.split('\n');
    const head = lines.slice(0, PREVIEW_LINES).join('\n');
    return lines.length > PREVIEW_LINES
      ? `${head}\n… (${lines.length - PREVIEW_LINES} more lines)`
      : head;
  });

  constructor(private readonly converter: ConverterService) {}

  protected onFileInput(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) void this.handle(file);
    input.value = '';
  }

  protected onDragOver(event: DragEvent): void {
    event.preventDefault();
    this.dragging.set(true);
  }

  protected onDragLeave(): void {
    this.dragging.set(false);
  }

  protected onDrop(event: DragEvent): void {
    event.preventDefault();
    this.dragging.set(false);
    const file = event.dataTransfer?.files?.[0];
    if (file) void this.handle(file);
  }

  protected download(): void {
    const current = this.result();
    if (!current) return;
    const blob = new Blob([current.xml], { type: 'application/xml' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = current.filename;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  protected async copy(): Promise<void> {
    const xml = this.result()?.xml;
    if (!xml) return;
    await navigator.clipboard.writeText(xml);
    this.copied.set(true);
    setTimeout(() => this.copied.set(false), 1500);
  }

  protected reset(): void {
    this.status.set('idle');
    this.result.set(null);
    this.errorMessage.set('');
    this.previewOpen.set(false);
  }

  private async handle(file: File): Promise<void> {
    this.status.set('parsing');
    this.result.set(null);
    this.errorMessage.set('');
    try {
      this.result.set(await this.converter.convert(file));
      this.status.set('done');
    } catch (err) {
      this.errorMessage.set(
        err instanceof ConverterError ? err.message : 'Something went wrong during conversion.',
      );
      this.status.set('error');
    }
  }
}
