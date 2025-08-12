/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      // Example: extend custom blur sizes or colors if needed
      blur: {
        xl: '12px',
      },
    },
  },
  plugins: [],
};
