import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import AgenteProdutividade from './AgenteProdutividade.jsx'

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <AgenteProdutividade />
  </StrictMode>,
)
