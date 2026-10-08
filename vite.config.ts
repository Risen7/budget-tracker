// Add Vite's React integration, including the React Compiler preset.
import react, { reactCompilerPreset } from '@vitejs/plugin-react'
// Load the Babel plugin used to configure the React Compiler.
import babel from '@rolldown/plugin-babel'
// Import Vite's helper for defining a typed configuration object.
import { defineConfig } from 'vite'

// https://vite.dev/config/
// Configure local development and the plugins that transform the application.
export default defineConfig({
  // Forward API requests to the separate Node server during development.
  server: {
    // Configure the proxy section for this project.
    proxy: {
      // Perform this step as part of the surrounding operation.
      '/api': 'http://localhost:3001'
    // Close the current object, callback, or expression.
    }
  // Close the current object, callback, or expression.
  },
  // Enable React support and compile components with the React Compiler.
  plugins: [
    // Transform JSX and apply the project's React plugin defaults.
    react(),
    // Add the React Compiler Babel preset to Vite's transformation pipeline.
    babel({ presets: [reactCompilerPreset()] })
  // Continue the surrounding list or object with this value.
  ],
// Close the current object, callback, or expression.
})
