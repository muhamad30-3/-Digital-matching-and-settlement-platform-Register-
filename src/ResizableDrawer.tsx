import React, { useState, useRef, useEffect } from 'react';
import { ChevronLeft, ChevronRight, GripVertical } from 'lucide-react';

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
}

export const ResizableDrawer: React.FC<ResizableDrawerProps> = ({
  title,
  icon = '💼',
  items,
  onSwap,
  onItemClick,
  minWidth = 300,
  maxWidth = 500,
  position = 'left',
}) => {
  const [isOpen, setIsOpen] = useState(true);
  const [width, setWidth] = useState(360);
  const [isDraggingResize, setIsDraggingResize] = useState(false);
  const [draggedItem, setDraggedItem] = useState<DrawerItem | null>(null);
  const [dragOverId, setDragOverId] = useState<string | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const resizeHandleRef = useRef<HTMLDivElement>(null);

  // معالج تغيير حجم القائمة
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

  // معالجات السحب والإفلات
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
        className={`group fixed top-1/2 -translate-y-1/2 z-40 flex h-14 w-8 items-center justify-center rounded-r-lg bg-gradient-to-b from-blue-500 to-blue-600 text-white hover:from-blue-600 hover:to-blue-700 shadow-lg transition-all duration-200`}
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
      className={`relative flex flex-col h-full bg-gradient-to-b from-slate-50 to-white shadow-2xl transition-all duration-200 border ${
        position === 'left'
          ? 'border-r border-slate-200'
          : 'border-l border-slate-200'
      }`}
      style={{
        width: `${width}px`,
      }}
      dir="rtl"
    >
      {/* رأس القائمة */}
      <div className="sticky top-0 z-10 border-b border-slate-200 bg-gradient-to-b from-blue-50 to-blue-100 px-4 py-4 shadow-sm">
        <div className="flex items-center justify-between gap-2">
          <div className="flex-1 min-w-0">
            <div className="flex items-center gap-2">
              <span className="text-lg">{icon}</span>
              <h2 className="font-bold text-slate-900 text-sm truncate">{title}</h2>
            </div>
            <p className="text-xs text-slate-600 mt-1">{items.length} عنصر</p>
          </div>
          <button
            onClick={() => setIsOpen(false)}
            className="rounded p-1.5 hover:bg-blue-200 transition-colors text-slate-600 shrink-0"
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
            <p className="text-center text-xs text-slate-400">
              اسحب العناصر هنا<br />أو استخدم الزر في الجدول
            </p>
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
                    ? 'opacity-50 scale-95 ring-2 ring-blue-500'
                    : dragOverId === item.id
                    ? 'border-green-400 bg-green-50 shadow-lg scale-105'
                    : 'border-slate-200 bg-white hover:border-blue-300 hover:shadow-md'
                }`}
              >
                {/* أيقونة السحب */}
                <div className="absolute right-2 top-2 opacity-0 group-hover:opacity-100 transition-opacity">
                  <GripVertical className="h-4 w-4 text-slate-400" />
                </div>

                <div className="space-y-2 pl-2">
                  {/* اسم الحوالة */}
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

                  {/* المبلغ والمصدر */}
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs text-slate-600">{item.source}</span>
                    <span className="text-sm font-bold text-blue-700 font-mono">
                      {item.amount.toLocaleString('ar-SA', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2,
                      })}
                    </span>
                  </div>

                  {/* التاريخ ورقم الفاتورة */}
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
      <div className="border-t border-slate-200 bg-slate-50 px-4 py-2 text-[10px] text-slate-500 font-medium">
        اسحب المربع على عنصر آخر لتبديلهما
      </div>

      {/* مقبض تغيير الحجم */}
      <div
        ref={resizeHandleRef}
        onMouseDown={() => setIsDraggingResize(true)}
        className={`absolute top-0 h-full w-1.5 cursor-col-resize select-none bg-gradient-to-b from-transparent via-blue-400 to-transparent opacity-0 hover:opacity-100 transition-opacity hover:bg-blue-500 ${
          position === 'left' ? '-right-0.75' : '-left-0.75'
        }`}
        title="اسحب لتغيير الحجم"
      />
    </div>
  );
};
