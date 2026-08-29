import { render } from '@testing-library/react'
import { describe, expect, it } from 'vitest'
import App from './App'

describe('App', () => {
  it('renders the popup shell', () => {
    render(<App />)
    expect(document.querySelector('.app-wrapper')).toBeInTheDocument()
  })
})
