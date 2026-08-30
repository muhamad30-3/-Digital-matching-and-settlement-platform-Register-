import React, { useState, useMemo } from 'react';
import { ResizableDrawer, DrawerItem } from './ResizableDrawer';

export interface Transfer {
  id: string;
  name: string;
  amount: number;
  source: 'cashier' | 'bank';
  date: string;
  invoiceId?: string;
  matchedWith?: string;
}

interface TransfersSectionProps {
  transfers: Transfer[];
  onSwap: (fromId: string, toId: string) => void;
  onTransferClick?: (transfer: Transfer) => void;
}

/**
 * قسم المطابقة المتقدم
 * يعرض قائمتين جانبيتين قابلتين للتوسيع
 * تدعم السحب والإفلات والتبديل الذكي
 */
export const TransfersSection: React.FC<TransfersSectionProps> = ({
  transfers,
  onSwap,
  onTransferClick,
}) => {
  // تقسيم الحوالات حسب المصدر
  const cashierItems = useMemo<DrawerItem[]>(
    () =>
      transfers
        .filter((t) => t.source === 'cashier')
        .map((t) => ({
          id: t.id,
          name: t.name,
          amount: t.amount,
          source: t.source,
          date: t.date,
          invoiceId: t.invoiceId,
          matchedWith: t.matchedWith,
        })),
    [transfers]
  );

  const bankItems = useMemo<DrawerItem[]>(
    () =>
      transfers
        .filter((t) => t.source === 'bank')
        .map((t) => ({
          id: t.id,
          name: t.name,
          amount: t.amount,
          source: t.source,
          date: t.date,
          invoiceId: t.invoiceId,
          matchedWith: t.matchedWith,
        })),
    [transfers]
  );

  const handleSwap = (fromId: string, toId: string) => {
    onSwap(fromId, toId);
  };

  return (
    <div className="relative w-full h-screen flex items-stretch" dir="rtl">
      {/* قائمة الكاشير اليسرى */}
      <div className="flex-shrink-0">
        <ResizableDrawer
          title="فواتير الكاشير"
          icon="💼"
          items={cashierItems}
          onSwap={handleSwap}
          onItemClick={(item) => {
            const transfer = transfers.find((t) => t.id === item.id);
            if (transfer) onTransferClick?.(transfer);
          }}
          minWidth={280}
          maxWidth={600}
          position="left"
        />
      </div>

      {/* منطقة المحتوى الوسطى */}
      <div className="flex-1 bg-gradient-to-br from-slate-50 to-white overflow-auto">
        <div className="p-6 space-y-6">
          <div className="rounded-xl border border-blue-200 bg-blue-50 p-4">
            <h2 className="text-2xl font-bold text-slate-900 mb-2">منطقة المطابقة الذكية</h2>
            <p className="text-slate-600 text-sm">
              اسحب الحوالات من القائمتين الجانبيتين وأفلتها على بعضها لمطابقتها تلقائياً
            </p>
          </div>

          {/* إحصائيات سريعة */}
          <div className="grid grid-cols-3 gap-4">
            <div className="p-4 rounded-xl bg-blue-50 border border-blue-200 shadow-sm hover:shadow-md transition-shadow">
              <p className="text-xs text-blue-600 font-bold mb-1">💼 فواتير الكاشير</p>
              <p className="text-3xl font-bold text-blue-900">{cashierItems.length}</p>
              <p className="text-xs text-blue-700 mt-2">عنصر جاهز</p>
            </div>
            <div className="p-4 rounded-xl bg-green-50 border border-green-200 shadow-sm hover:shadow-md transition-shadow">
              <p className="text-xs text-green-600 font-bold mb-1">🏦 حوالات البنك</p>
              <p className="text-3xl font-bold text-green-900">{bankItems.length}</p>
              <p className="text-xs text-green-700 mt-2">عنصر جاهز</p>
            </div>
            <div className="p-4 rounded-xl bg-amber-50 border border-amber-200 shadow-sm hover:shadow-md transition-shadow">
              <p className="text-xs text-amber-600 font-bold mb-1">✅ مطابقات</p>
              <p className="text-3xl font-bold text-amber-900">
                {transfers.filter((t) => t.matchedWith).length}
              </p>
              <p className="text-xs text-amber-700 mt-2">مطابقة ناجحة</p>
            </div>
          </div>

          {/* قائمة الحوالات المطابقة */}
          <div className="mt-8 space-y-3">
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <span className="text-2xl">✅</span>
              الحوالات المطابقة
            </h3>
            <div className="space-y-2 max-h-96 overflow-y-auto">
              {transfers
                .filter((t) => t.matchedWith)
                .map((transfer) => (
                  <div
                    key={transfer.id}
                    className="p-4 rounded-lg border-2 border-green-200 bg-green-50 hover:shadow-md transition-shadow"
                  >
                    <div className="flex items-center justify-between gap-4">
                      <div className="flex-1 min-w-0">
                        <p className="font-semibold text-slate-900 truncate">{transfer.name}</p>
                        <p className="text-xs text-slate-500 mt-1">📅 {transfer.date}</p>
                      </div>
                      <div className="text-right shrink-0">
                        <p className="font-bold text-green-700 font-mono text-lg">
                          {transfer.amount.toLocaleString('ar-SA', {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2,
                          })}
                        </p>
                        <span className="inline-block text-xs px-2.5 py-1 rounded-full bg-green-200 text-green-800 font-bold mt-1">
                          ✓ مطابقة
                        </span>
                      </div>
                    </div>
                  </div>
                ))}
              {transfers.filter((t) => t.matchedWith).length === 0 && (
                <div className="flex flex-col items-center justify-center py-12 text-center">
                  <p className="text-4xl mb-2">🔄</p>
                  <p className="text-slate-500 font-medium">ابدأ بسحب الحوالات من القائمتين</p>
                  <p className="text-slate-400 text-sm mt-1">لا توجد مطابقات حتى الآن</p>
                </div>
              )}
            </div>
          </div>

          {/* نصائح الاستخدام */}
          <div className="rounded-lg bg-indigo-50 border border-indigo-200 p-4 text-sm text-indigo-800">
            <p className="font-bold mb-2">💡 نصائح الاستخدام:</p>
            <ul className="space-y-1 text-xs">
              <li>✓ اسحب مربع من القائمة اليسرى وأفلته على مربع في القائمة اليمنى</li>
              <li>✓ يمكنك توسيع أو تصغير القائمتين بسحب الحافة</li>
              <li>✓ المطابقة تتم تلقائياً عند الإفلات</li>
              <li>✓ الحوالات المطابقة ستظهر في القائمة أعلاه</li>
            </ul>
          </div>
        </div>
      </div>

      {/* قائمة البنك اليمنى */}
      <div className="flex-shrink-0">
        <ResizableDrawer
          title="حوالات البنك"
          icon="🏦"
          items={bankItems}
          onSwap={handleSwap}
          onItemClick={(item) => {
            const transfer = transfers.find((t) => t.id === item.id);
            if (transfer) onTransferClick?.(transfer);
          }}
          minWidth={280}
          maxWidth={600}
          position="right"
        />
      </div>
    </div>
  );
};
