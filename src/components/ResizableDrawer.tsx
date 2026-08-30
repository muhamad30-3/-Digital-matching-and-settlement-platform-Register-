import React, { useState, useRef, useEffect } from 'react';
import { ChevronLeft, ChevronRight } from 'lucide-react';

export interface DrawerTransfer {
  id: string;
  amount: number;
  source: string;
  date: string;
  invoiceId?: string;
  matchedWith?: string;
}

interface ResizableDrawerProps {
  title: string;
  items: DrawerTransfer[];
  onSwap: (draggedId: string, droppedId: string, fromDrawer: string, toDrawer: string) => void;
  onItemSelect?: (item: DrawerTransfer) => void;
  minWidth?: number;
  maxWidth?: number;
  position?: 'left' | 'right';
  drawerId: string;
}

export const ResizableDrawer: React.FC<ResizableDrawerProps> = ({
  title,
  items,
  onSwap,
  onItemSelect,
  minWidth = 280,
  maxWidth = 600,
  position = 'left',
  drawerId,
}) => {
  const [isOpen, setIsOpen] = useState(true);
  const [width, setWidth] = useState(360);
  const [isDraggingResize, setIsDraggingResize] = useState(false);
  const [draggedItem, setDraggedItem] = useState<DrawerTransfer | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const resizeRef = useRef<HTMLDivElement>(null);

  // Resize functionality
  useEffect(() => {
    if (!isDraggingResize) return;

    const handleMouseMove = (e: MouseEvent) => {
      if (!containerRef.current) return;

      const rect = containerRef.current.parentElement?.getBoundingClientRect();
      if (!rect) return;

      let newWidth: number;
      if (position === 'left') {
        newWidth = e.clientX - rect.left;
      } else {
        newWidth = rect.right - e.clientX;
      }

      newWidth = Math.max(minWidth, Math.min(maxWidth, newWidth));
      setWidth(newWidth);
    };

    const handleMouseUp = () => {
      setIsDraggingResize(false);
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);

    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isDraggingResize, minWidth, maxWidth, position]);

  // Drag and drop handlers
  const handleDragStart = (item: DrawerTransfer) => {
    setDraggedItem(item);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
  };

  const handleDrop = (targetItem: DrawerTransfer | null) => {
    if (!draggedItem) return;

    if (targetItem && draggedItem.id !== targetItem.id) {
      onSwap(draggedItem.id, targetItem.id, draggedItem.matchedWith || 'source', drawerId);
    }

    setDraggedItem(null);
  };

  if (!isOpen) {
    return (
      <div
        className={`group relative flex ${position === 'left' ? 'flex-row-reverse' : ''}`}
      >
        <button
          onClick={() => setIsOpen(true)}
          className="absolute top-1/2 -translate-y-1/2 z-50 flex h-12 w-8 items-center justify-center rounded-r-lg bg-blue-600 text-white hover:bg-blue-700 transition-colors"
          style={{
            [position === 'left' ? 'left' : 'right']: 0,
          }}
          title={`فتح ${title}`}
        >
          {position === 'left' ? (
            <ChevronRight className="h-5 w-5" />
          ) : (
            <ChevronLeft className="h-5 w-5" />
          )}
        </button>
      </div>
    );
  }

  return (
    <div
      ref={containerRef}
      className={`relative flex flex-col h-full bg-white border-r border-slate-200 shadow-lg transition-all duration-200 ${
        position === 'right' ? 'border-r-0 border-l' : ''
      }`}
      style={{
        width: `${width}px`,
        [position]: 0,
      }}
      dir="rtl"
    >
      {/* Header */}
      <div className="flex items-center justify-between border-b border-slate-200 bg-gradient-to-b from-blue-50 to-white px-4 py-3">
        <h2 className="font-bold text-slate-900 text-sm">{title}</h2>
        <div className="flex items-center gap-1">
          <button
            onClick={() => setIsOpen(false)}
            className="rounded p-1 hover:bg-slate-200 transition-colors text-slate-600"
            title="إغلاق"
          >
            {position === 'left' ? (
              <ChevronLeft className="h-4 w-4" />
            ) : (
              <ChevronRight className="h-4 w-4" />
            )}
          </button>
        </div>
      </div>

      {/* Items Container */}
      <div className="flex-1 overflow-y-auto">
        {items.length === 0 ? (
          <div className="flex h-full items-center justify-center">
            <p className="text-xs text-slate-400">لا توجد عناصر</p>
          </div>
        ) : (
          <div className="space-y-2 p-3">
            {items.map((item) => (
              <div
                key={item.id}
                draggable
                onDragStart={() => handleDragStart(item)}
                onDragOver={handleDragOver}
                onDrop={() => handleDrop(item)}
                onClick={() => onItemSelect?.(item)}
                className={`group cursor-move rounded-lg border-2 border-slate-200 bg-white p-3 transition-all hover:border-blue-400 hover:shadow-md ${
                  draggedItem?.id === item.id
                    ? 'opacity-50 ring-2 ring-blue-500'
                    : ''
                }`}
              >
                <div className="space-y-1">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-slate-700">
                      {item.source}
                    </span>
                    <span className="text-xs text-slate-500 font-mono">
                      {item.id.slice(0, 6)}
                    </span>
                  </div>
                  <div className="text-sm font-bold text-green-700">
                    {item.amount.toFixed(2)} ريال
                  </div>
                  <div className="flex items-center justify-between text-[10px] text-slate-500">
                    <span>{item.date}</span>
                    {item.matchedWith && (
                      <span className="rounded bg-amber-100 px-1.5 py-0.5 text-amber-700">
                        مطابقة
                      </span>
                    )}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Resize Handle */}
      <div
        ref={resizeRef}
        onMouseDown={() => setIsDraggingResize(true)}
        className={`absolute top-0 h-full w-1 cursor-col-resize select-none bg-gradient-to-b from-transparent via-blue-400 to-transparent opacity-0 hover:opacity-100 transition-opacity ${
          position === 'left' ? '-right-0.5' : '-left-0.5'
        }`}
        title="اسحب للتوسيع/التصغير"
      />

      {/* Footer Info */}
      <div className="border-t border-slate-200 bg-slate-50 px-4 py-2 text-[11px] text-slate-500">
        {items.length} عنصر
      </div>
    </div>
  );
};
