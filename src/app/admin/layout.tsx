"use client"

import { useEffect, useState, Suspense } from 'react'
import { useRouter, usePathname } from 'next/navigation'
import { Sidebar } from './components/sidebar'

export default function AdminLayout({
  children,
}: {
  children: React.ReactNode
}) {
  const router = useRouter()
  const pathname = usePathname()
  const [token, setToken] = useState<string | null>(null)
  const [user, setUser] = useState<{ email: string; role: string } | null>(null)
  const isLoginPage = pathname === '/admin/login'

  useEffect(() => {
    if (isLoginPage) return

    const storedToken = localStorage.getItem('adminToken')
    if (storedToken) {
      setToken(storedToken)
      try {
        const payload = JSON.parse(atob(storedToken.split('.')[1]))
        setUser({ email: payload.email, role: payload.role })
      } catch (e) {
        console.error('Error decoding token:', e)
      }
    } else {
      router.push('/admin/login')
    }
  }, [router, isLoginPage])

  const handleLogout = () => {
    localStorage.removeItem('adminToken')
    router.push('/admin/login')
  }

  if (isLoginPage) {
    return <>{children}</>
  }

  if (!token) {
    return (
      <div className="flex items-center justify-center min-h-screen bg-background">
        <div className="flex flex-col items-center gap-3">
          <div className="w-8 h-8 border-2 border-primary border-t-transparent rounded-full animate-spin" />
          <p className="text-sm text-muted-foreground">Loading...</p>
        </div>
      </div>
    )
  }

  return (
    <div className="flex min-h-screen bg-background">
      <Suspense>
        <Sidebar user={user} onLogout={handleLogout} />
      </Suspense>
      <main className="flex-1 overflow-auto">
        {children}
      </main>
    </div>
  )
}
