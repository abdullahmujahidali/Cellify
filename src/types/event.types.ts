/**
 * Event types for Cellify
 */

import type { CellValue } from './cell.types.js';
import type { CellStyle } from './style.types.js';

/**
 * Base event interface
 */
export interface SheetEvent {
  /** The sheet where the event occurred */
  sheetName: string;
  /** Timestamp of the event */
  timestamp: number;
}

/**
 * Cell change event
 */
export interface CellChangeEvent extends SheetEvent {
  type: 'cellChange';
  /** Cell address in A1 notation */
  address: string;
  /** Row index (0-based) */
  row: number;
  /** Column index (0-based) */
  col: number;
  /** Previous value */
  oldValue: CellValue;
  /** New value */
  newValue: CellValue;
}

/**
 * Cell style change event
 */
export interface CellStyleChangeEvent extends SheetEvent {
  type: 'cellStyleChange';
  /** Cell address in A1 notation */
  address: string;
  /** Row index (0-based) */
  row: number;
  /** Column index (0-based) */
  col: number;
  /** Previous style */
  oldStyle: CellStyle | undefined;
  /** New style */
  newStyle: CellStyle | undefined;
}

/**
 * Range change event (for batch operations)
 */
export interface RangeChangeEvent extends SheetEvent {
  type: 'rangeChange';
  /** Range in A1 notation (e.g., "A1:C3") */
  range: string;
  /** Start row */
  startRow: number;
  /** Start column */
  startCol: number;
  /** End row */
  endRow: number;
  /** End column */
  endCol: number;
  /** Number of cells affected */
  cellCount: number;
}

/**
 * Cell added event
 */
export interface CellAddedEvent extends SheetEvent {
  type: 'cellAdded';
  /** Cell address in A1 notation */
  address: string;
  /** Row index (0-based) */
  row: number;
  /** Column index (0-based) */
  col: number;
}

/**
 * Cell deleted event
 */
export interface CellDeletedEvent extends SheetEvent {
  type: 'cellDeleted';
  /** Cell address in A1 notation */
  address: string;
  /** Row index (0-based) */
  row: number;
  /** Column index (0-based) */
  col: number;
  /** Value at time of deletion */
  value: CellValue;
}

/**
 * Union of all event types
 */
export type SheetEventType =
  | CellChangeEvent
  | CellStyleChangeEvent
  | RangeChangeEvent
  | CellAddedEvent
  | CellDeletedEvent;

/**
 * Event handler function
 */
export type SheetEventHandler<T extends SheetEventType = SheetEventType> = (event: T) => void;

/**
 * Map of event types to their handlers
 */
export interface SheetEventMap {
  cellChange: CellChangeEvent;
  cellStyleChange: CellStyleChangeEvent;
  rangeChange: RangeChangeEvent;
  cellAdded: CellAddedEvent;
  cellDeleted: CellDeletedEvent;
  '*': SheetEventType; // Wildcard for all events
}

/**
 * Change record for tracking modifications
 */
export interface ChangeRecord {
  /** Unique ID for this change */
  id: string;
  /** Type of change */
  type: 'value' | 'style' | 'formula' | 'delete';
  /** Cell address */
  address: string;
  /** Row index */
  row: number;
  /** Column index */
  col: number;
  /** Value before change */
  oldValue?: CellValue;
  /** Value after change */
  newValue?: CellValue;
  /** Style before change */
  oldStyle?: CellStyle;
  /** Style after change */
  newStyle?: CellStyle;
  /** Timestamp */
  timestamp: number;
}
