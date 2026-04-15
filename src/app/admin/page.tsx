"use client"

import { useEffect, useState, Suspense } from 'react'
import { useRouter, useSearchParams } from 'next/navigation'
import { UrlTable } from './components/url-table'
import { AuditTable } from './components/audit-table'

function AdminContent() {
  const router = useRouter()
  const searchParams = useSearchParams()
  const [token, setToken] = useState<string | null>(null)
  const [mounted, setMounted] = useState(false)

  const activeTab = searchParams.get('tab') ?? 'urls'

  useEffect(() => {
    setMounted(true)
    const storedToken = localStorage.getItem('adminToken')
    if (!storedToken) {
      router.push('/admin/login')
    } else {
      setToken(storedToken)
    }
  }, [router])

  if (!mounted || !token) {
    return (
      <div className="flex items-center justify-center h-64">
        <div className="w-8 h-8 border-2 border-primary border-t-transparent rounded-full animate-spin" />
      </div>
    )
  }

  const titles: Record<string, string> = {
    urls: 'URL Checks',
    audit: 'Audit Records',
  }

  return (
    <div className="p-6 space-y-6">
      <div>
        <h2 className="text-xl font-semibold text-foreground">{titles[activeTab] ?? 'Dashboard'}</h2>
        <p className="text-sm text-muted-foreground mt-1">
          {activeTab === 'urls'
            ? 'All URL scans performed through PhishGuardAI'
            : 'User actions and access audit trail'}
        </p>
      </div>
      {activeTab === 'urls' && <UrlTable token={token} />}
      {activeTab === 'audit' && <AuditTable token={token} />}
    </div>
  )
}

export default function AdminPage() {
  return (
    <Suspense>
      <AdminContent />
    </Suspense>
  )
}
