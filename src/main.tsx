// Import React's development checks so invalid component usage is easier to catch.
import { StrictMode } from 'react'
// Import the browser API that attaches a React tree to an HTML element.
import { createRoot } from 'react-dom/client'
// Load global styles before rendering the application.
import './index.css'
// Import the main budget-tracker dashboard component.
import App from './App.tsx'

// Find the page's root container, create the React root, and render the dashboard in strict mode.
createRoot(document.getElementById('root')!).render(
  // StrictMode helps expose unsafe React patterns during development.
  <StrictMode>
    {/* Render the entire budget-tracking interface. */}
    <App />
  </StrictMode>,
)
