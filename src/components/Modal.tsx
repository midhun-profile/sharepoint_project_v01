import type { ReactNode } from 'react';

interface ModalProps {
  isOpen?: boolean;
  onClose?: () => void;
  title?: string;
  children?: ReactNode;
}

export function Modal({ isOpen, onClose, title, children }: ModalProps) {
  if (!isOpen) return null;

  return (
    <div id="modal-backdrop" className="fixed inset-0 z-50 flex items-center justify-center bg-black/50">
      <div id="modal-container" className="bg-white rounded-lg p-6 max-w-md w-full shadow-lg">
        {title && <h3 className="text-lg font-medium mb-4">{title}</h3>}
        <div>{children}</div>
        <div className="mt-6 flex justify-end">
          <button
            id="modal-close-btn"
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-sm bg-neutral-100 hover:bg-neutral-200 rounded"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
}
