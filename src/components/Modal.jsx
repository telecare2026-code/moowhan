import useEscapeKey from '../hooks/useEscapeKey.js';

// Full-screen overlay. Closes on Escape and on backdrop click.
export default function Modal({ onClose, className = '', children, label }) {
  useEscapeKey(onClose);
  return (
    <div
      className="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 flex items-center justify-center p-2 sm:p-4"
      onMouseDown={(e) => { if (e.target === e.currentTarget) onClose?.(); }}
      role="dialog"
      aria-modal="true"
      aria-label={label}
    >
      <div className={`bg-white rounded-2xl shadow-2xl w-full flex flex-col overflow-hidden ${className}`}>
        {children}
      </div>
    </div>
  );
}
