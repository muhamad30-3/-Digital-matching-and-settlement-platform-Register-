import React, { useState, useRef, useEffect } from 'react';
import { ChevronLeft, ChevronRight, GripVertical, AlertCircle } from 'lucide-react';

export interface DrawerItem {
  id: string;
  name: string;
  amount: number;
  source: string;
  date: string;
  matchedWith?: string;
  invoiceId?: string;
}

interface ResizableDrawerProps {
  title: string;
  icon?: string;
  items: DrawerItem[];
  onSwap: (draggedId: string, droppedId: string) => void;
  onItemClick?: (item: DrawerItem) => void;
  minWidth?: number;
  maxWidth?: number;
  position?: 'left' | 'right';
  accentColor?: 'blue' | 'green' | 'amber' | 'purple' | 'red';
}

const accentColors = {
  blue: {
    bg: 'bg-blue-50',
    border: 'border-blue-200',
    header: 'from-blue-50 to-blue-100',
    button: 'from-blue-500 to-blue-600 hover:from-blue-600 hover:to-blue-700',
    text: 'text-blue-700',
    textDark: 'text-blue-900',
  },
  green: {
    bg: 'bg-green-50',
    border: 'border-green-200',
    header: 'from-green-50 to-green-100',
    button: 'from-green-500 to-green-600 hover:from-green-600 hover:to-green-700',
    text: 'text-green-700',
    textDark: 'text-green-900',
  },
  amber: {
    bg: 'bg-amber-50',
    border: 'border-amber-200',
    header: 'from-amber-50 to-amber-100',
    button: 'from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700',
    text: 'text-amber-700',
    textDark: 'text-amber-900',
  },
  purple: {
    bg: 'bg-purple-50',
    border: 'border-purple-200',
    header: 'from-purple-50 to-purple-100',
    button: 'from-purple-500 to-purple-600 hover:from-purple-600 hover:to-purple-700',
    text: 'text-purple-700',
    textDark: 'text-purple-900',
  },
  red: {
    bg: 'bg-red-50',
    border: 'border-red-200',
    header: 'from-red-50 to-red-100',
    button: 'from-red-500 to-red-600 hover:from-red-600 hover:to-red-700',
    text: 'text-red-700',
    textDark: 'text-red-900',
  },
};

export const ResizableDrawer: React.FC<ResizableDrawerProps> = ({
  title,
  icon = '💼',
  items,
  onSwap,
  onItemClick,
  minWidth = 300,
  maxWidth = 500,
  position = 'left',
  accentColor = 'blue',
}) => {
  const [isOpen, setIsOpen] = useState(true);
  const [width, setWidth] = useState(360);
  const [isDraggingResize, setIsDraggingResize] = useState(false);
  const [draggedItem, setDraggedItem] = useState<DrawerItem | null>(null);
  const [dragOverId, setDragOverId] = useState<string | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const resizeHandleRef = useRef<HTMLDivElement>(null);

  const colors = accentColors[accentColor];

  useEffect(() => {
    if (!isDraggingResize) return;

    const handleMouseMove = (e: MouseEvent) => {
      if (!containerRef.current) return;
      const container = containerRef.current;
      const rect = container.getBoundingClientRect();
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

  const handleDragStart = (e: React.DragEvent, item: DrawerItem) => {
    setDraggedItem(item);
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('application/json', JSON.stringify(item));
  };

  const handleDragOver = (e: React.DragEvent, item: DrawerItem) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setDragOverId(item.id);
  };

  const handleDragLeave = () => {
    setDragOverId(null);
  };

  const handleDrop = (e: React.DragEvent, targetItem: DrawerItem) => {
    e.preventDefault();
    e.stopPropagation();
    if (draggedItem && draggedItem.id !== targetItem.id) {
      onSwap(draggedItem.id, targetItem.id);
    }
    setDraggedItem(null);
    setDragOverId(null);
  };

  const handleDragEnd = () => {
    setDraggedItem(null);
    setDragOverId(null);
  };

  if (!isOpen) {
    return (
      <button
        onClick={() => setIsOpen(true)}
        className={`group fixed top-1/2 -translate-y-1/2 z-40 flex h-14 w-10 items-center justify-center rounded-r-lg bg-gradient-to-b ${colors.button} text-white shadow-lg transition-all duration-200`}
        style={{
          [position === 'left' ? 'left' : 'right']: 0,
        }}
        title={`فتح ${title}`}
      >
        <div className="flex items-center gap-1">
          {position === 'left' ? (
            <ChevronRight className="h-5 w-5 group-hover:translate-x-0.5 transition-transform" />
          ) : (
            <ChevronLeft className="h-5 w-5 group-hover:-translate-x-0.5 transition-transform" />
          )}
          <span className="text-xs font-bold">{items.length}</span>
        </div>
      </button>
    );
  }

  return (
    <div
      ref={containerRef}
      className={`relative flex flex-col h-full ${colors.bg} shadow-2xl transition-all duration-200 border ${
        position === 'left' ? `border-r ${colors.border}` : `border-l ${colors.border}`
      }`}
      style={{
        width: `${width}px`,
      }}
      dir="rtl"
    >
      {/* رأس القائمة */}
      <div className={`sticky top-0 z-10 border-b ${colors.border} bg-gradient-to-b ${colors.header} px-4 py-4 shadow-sm`}>
        <div className="flex items-center justify-between gap-2">
          <div className="flex-1 min-w-0">
            <div className="flex items-center gap-2">
              <span className="text-lg">{icon}</span>
              <h2 className={`font-bold text-slate-900 text-sm truncate`}>{title}</h2>
            </div>
            <p className={`text-xs ${colors.text} mt-1`}>{items.length} عنصر</p>
          </div>
          <button
            onClick={() => setIsOpen(false)}
            className={`rounded p-1.5 transition-colors text-slate-600 shrink-0 hover:${colors.bg}`}
            title="إغلاق"
          >
            {position === 'left' ? (
              <ChevronLeft className="h-5 w-5" />
            ) : (
              <ChevronRight className="h-5 w-5" />
            )}
          </button>
        </div>
      </div>

      {/* قائمة العناصر */}
      <div className="flex-1 overflow-y-auto">
        {items.length === 0 ? (
          <div className="flex h-full items-center justify-center p-4">
            <div className="text-center">
              <AlertCircle className={`h-8 w-8 ${colors.text} mx-auto mb-2 opacity-50`} />
              <p className="text-xs text-slate-400">اسحب العناصر هنا</p>
            </div>
          </div>
        ) : (
          <div className="space-y-2 p-3">
            {items.map((item) => (
              <div
                key={item.id}
                draggable
                onDragStart={(e) => handleDragStart(e, item)}
                onDragOver={(e) => handleDragOver(e, item)}
                onDragLeave={handleDragLeave}
                onDrop={(e) => handleDrop(e, item)}
                onDragEnd={handleDragEnd}
                onClick={() => onItemClick?.(item)}
                className={`group relative cursor-move rounded-lg border-2 transition-all duration-200 p-3 ${
                  draggedItem?.id === item.id
                    ? `opacity-50 scale-95 ring-2 ${colors.text.replace('text-', 'ring-')}`
                    : dragOverId === item.id
                    ? `border-green-400 bg-green-50 shadow-lg scale-105`
                    : `border-slate-200 bg-white hover:border-slate-300 hover:shadow-md`
                }`}
              >
                <div className="absolute right-2 top-2 opacity-0 group-hover:opacity-100 transition-opacity">
                  <GripVertical className="h-4 w-4 text-slate-400" />
                </div>

                <div className="space-y-2 pl-2">
                  <div className="flex items-start justify-between gap-2">
                    <span className="text-sm font-semibold text-slate-800 truncate flex-1">
                      {item.name}
                    </span>
                    {item.matchedWith && (
                      <span className="text-[10px] px-1.5 py-0.5 rounded bg-green-100 text-green-700 border border-green-300 font-medium shrink-0">
                        ✓ مطابقة
                      </span>
                    )}
                  </div>

                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs text-slate-600">{item.source}</span>
                    <span className={`text-sm font-bold ${colors.text} font-mono`}>
                      {item.amount.toLocaleString('ar-SA', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2,
                      })}
                    </span>
                  </div>

                  <div className="flex items-center justify-between text-[10px] text-slate-500">
                    <span>{item.date}</span>
                    {item.invoiceId && (
                      <span className="font-mono bg-slate-100 px-1 py-0.5 rounded">
                        #{item.invoiceId}
                      </span>
                    )}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* التذييل */}
      <div className={`border-t ${colors.border} bg-slate-50 px-4 py-2 text-[10px] text-slate-500 font-medium`}>
        اسحب على عنصر آخر لتبديله
      </div>

      {/* مقبض تغيير الحجم */}
      <div
        ref={resizeHandleRef}
        onMouseDown={() => setIsDraggingResize(true)}
        className={`absolute top-0 h-full w-1.5 cursor-col-resize select-none bg-gradient-to-b from-transparent via-slate-300 to-transparent opacity-0 hover:opacity-100 transition-opacity hover:bg-slate-400 ${
          position === 'left' ? '-right-0.75' : '-left-0.75'
        }`}
        title="اسحب لتغيير الحجم"
      />
    </div>
  );
};
