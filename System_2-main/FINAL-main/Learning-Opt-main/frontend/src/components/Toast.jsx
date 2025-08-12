import { useEffect, useState } from 'react';

export default function Toast({ message, type = 'info' }) {
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    setVisible(true);
    const timer = setTimeout(() => setVisible(false), 3000);
    return () => clearTimeout(timer);
  }, [message]);

  return (
    <div
      className={`fixed bottom-6 left-1/2 transform -translate-x-1/2 px-6 py-3 rounded-lg text-white shadow-lg transition-opacity duration-300 z-50 ${
        visible ? 'opacity-100' : 'opacity-0'
      } ${
        type === 'success'
          ? 'bg-green-600'
          : type === 'error'
          ? 'bg-red-600'
          : 'bg-zinc-800'
      }`}
    >
      {message}
    </div>
  );
}
