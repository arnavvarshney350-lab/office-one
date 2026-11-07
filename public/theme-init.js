// Applies the saved (or system) theme before first paint so there is no flash of the wrong theme.
try {
  var t = localStorage.getItem('oo-theme')
  document.documentElement.dataset.theme = t || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light')
} catch { /* storage blocked: the CSS default (light) applies */ }
