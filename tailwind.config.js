/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./index.html",
    "./assets/js/**/*.js"
  ],
  theme: {
    extend: {
      colors: {
        primary: "#005bbf",
        "primary-container": "#1a73e8",
        "on-primary": "#ffffff",
        secondary: "#006e2c",
        "secondary-container": "#86f898",
        "on-secondary-container": "#00722f",
        tertiary: "#795900",
        "tertiary-fixed": "#ffdfa0",
        "tertiary-fixed-dim": "#fbbc05",
        error: "#ba1a1a",
        "error-container": "#ffdad6",
        "on-error-container": "#93000a",
        surface: "#f8f9fa",
        "surface-dim": "#d9dadb",
        "surface-bright": "#f8f9fa",
        "surface-container-lowest": "#ffffff",
        "surface-container-low": "#f3f4f5",
        "surface-container": "#edeeef",
        "surface-container-high": "#e7e8e9",
        "surface-container-highest": "#e1e3e4",
        "on-surface": "#191c1d",
        "on-surface-variant": "#414754",
        outline: "#727785",
        "outline-variant": "#c1c6d6"
      }
    }
  },
  plugins: [
    require('@tailwindcss/typography'),
    require('@tailwindcss/forms')
  ]
};
