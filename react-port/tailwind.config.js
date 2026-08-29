/** @type {import('tailwindcss').Config} */
export default {
  content: [
    './apps/*/index.html',
    './apps/*/src/**/*.{ts,tsx}',
    './packages/shared/src/**/*.{ts,tsx}',
  ],
  theme: {
    extend: {
      colors: {
        page: '#F8FAFC',
        card: '#FFFFFF',
        ink: {
          main: '#0F172A',
          muted: '#475569',
          dim: '#94A3B8',
        },
        teal: {
          DEFAULT: '#004E5A',
          hover: '#003B46',
        },
        yellow: {
          DEFAULT: '#FFD100',
          light: '#FACC15',
        },
        line: {
          subtle: '#E2E8F0',
          input: '#CBD5E1',
          inputHover: '#94A3B8',
        },
      },
      boxShadow: {
        card: '0 10px 30px -4px rgba(0, 0, 0, 0.08), 0 4px 10px -2px rgba(0, 0, 0, 0.04)',
        sm: '0 2px 6px 0 rgba(0, 0, 0, 0.05)',
        'yellow-glow': '0 4px 14px rgba(255, 209, 0, 0.35)',
        'yellow-glow-lg': '0 6px 18px rgba(255, 209, 0, 0.45)',
        'red-glow': '0 4px 14px rgba(239, 68, 68, 0.35)',
        'emerald-glow': '0 4px 14px rgba(16, 185, 129, 0.35)',
        'emerald-glow-lg': '0 6px 18px rgba(16, 185, 129, 0.45)',
        modal: '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
      },
      borderRadius: {
        sm: '14px',
        md: '18px',
        lg: '24px',
      },
      fontFamily: {
        sans: [
          'Inter',
          '-apple-system',
          'BlinkMacSystemFont',
          '"Segoe UI"',
          'Roboto',
          'sans-serif',
        ],
      },
    },
    keyframes: {
      'ping-slow': {
        '0%, 100%': { transform: 'scale(0.95)', opacity: '1' },
        '50%': { transform: 'scale(1.3)', opacity: '0.6' },
      },
      'fade-in': {
        from: { opacity: '0' },
        to: { opacity: '1' },
      },
      'scale-up': {
        from: { opacity: '0', transform: 'scale(0.92)' },
        to: { opacity: '1', transform: 'scale(1)' },
      },
      'slide-in': {
        from: { opacity: '0', transform: 'translateY(-6px)' },
        to: { opacity: '1', transform: 'translateY(0)' },
      },
    },
    animation: {
      'ping-slow': 'ping-slow 1.5s infinite',
      'fade-in': 'fade-in 0.3s ease',
      'scale-up': 'scale-up 0.2s cubic-bezier(0.34, 1.56, 0.64, 1)',
      'slide-in': 'slide-in 0.2s ease-out',
    },
  },
  plugins: [],
}
