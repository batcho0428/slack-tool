'use client';

type Props = {
  open: boolean;
  title: string;
  message: string;
  onClose: () => void;
};

export function Dialog({ open, title, message, onClose }: Props) {
  if (!open) return null;
  return (
    <div style={overlay}>
      <div style={panel}>
        <h3 style={{ marginTop: 0 }}>{title}</h3>
        <p style={{ whiteSpace: 'pre-wrap' }}>{message}</p>
        <div style={{ textAlign: 'right' }}>
          <button onClick={onClose} style={button}>閉じる</button>
        </div>
      </div>
    </div>
  );
}

const overlay: any = {
  position: 'fixed',
  inset: 0,
  background: 'rgba(0,0,0,0.45)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 100
};

const panel: any = {
  width: 'min(560px, calc(100% - 24px))',
  background: '#fff',
  borderRadius: 12,
  padding: 16,
  border: '1px solid #e5e7eb'
};

const button: any = {
  border: '1px solid #d1d5db',
  background: '#fff',
  borderRadius: 8,
  padding: '8px 14px',
  cursor: 'pointer'
};
