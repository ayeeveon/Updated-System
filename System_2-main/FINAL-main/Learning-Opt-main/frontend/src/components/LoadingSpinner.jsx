export default function LoadingSpinner({ size = 40, color = 'white', overlay = false }) {
  const spinner = (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      className="text-white"
      fill="none"
      viewBox="0 0 24 24"
      width={size}
      height={size}
      stroke={color}
      strokeWidth="4"
    >
      {/*Track Circle*/}
      <circle
        className="text-zinc-600"
        cx="12"
        cy="12"
        r="10"
        stroke="currentColor"
        strokeWidth="4"
        fill="none"
      />
      {/*Spinner*/}
      <circle
        cx="12"
        cy="12"
        r="10"
        stroke="currentColor"
        strokeWidth="4"
        strokeLinecap="round"
        fill="none"
        strokeDasharray="60"
        strokeDashoffset="15"
        className="animate-spin origin-center"
      />
    </svg>
  );

  if (!overlay) return spinner;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center pointer-events-none">
      <div className="bg-black opacity-60 rounded-xl p-6 shadow-lg flex items-center justify-center pointer-events-auto">
        {spinner}
      </div>
    </div>
  );
}
