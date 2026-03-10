import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'PhishGuardAI',
  description: 'Admin dashboard for PhishGuardAI',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}

