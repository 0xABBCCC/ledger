import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
// import './index.css'
// import App from './App.jsx'
import Ledger from './App.jsx'

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <Ledger />
  </StrictMode>,
)
