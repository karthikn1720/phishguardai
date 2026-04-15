"use client"

import Link from 'next/link'
import { useRouter, useSearchParams } from 'next/navigation'
import { ShieldCheck, Link2, ClipboardList, LogOut, User } from 'lucide-react'

interface SidebarProps {
  user: { email: string; role: string } | null
  onLogout: () => void
}

const navItems = [
  { label: 'URL Checks', tab: 'urls', icon: Link2 },
  { label: 'Audit Records', tab: 'audit', icon: ClipboardList },
]

export function Sidebar({ user, onLogout }: SidebarProps) {
  const searchParams = useSearchParams()
  const activeTab = searchParams.get('tab') ?? 'urls'

  return (
    <aside className="flex flex-col w-60 min-h-screen border-r border-border bg-sidebar shrink-0">
      {/* Brand */}
      <div className="flex items-center gap-2.5 px-5 py-5 border-b border-border">
        <div className="flex items-center justify-center w-8 h-8 rounded-lg bg-primary/20">
          <ShieldCheck className="w-5 h-5 text-primary" />
        </div>
        <div>
          <p className="text-sm font-bold text-sidebar-foreground leading-none">PhishGuardAI</p>
          <p className="text-xs text-muted-foreground mt-0.5">Admin Panel</p>
        </div>
      </div>

      {/* Nav */}
      <nav className="flex flex-col gap-1 px-3 py-4 flex-1">
        <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider px-2 mb-2">
          Dashboard
        </p>
        {navItems.map(({ label, tab, icon: Icon }) => {
          const isActive = activeTab === tab
          return (
            <Link
              key={tab}
              href={`/admin?tab=${tab}`}
              className={`flex items-center gap-3 px-3 py-2 rounded-md text-sm transition-colors ${
                isActive
                  ? 'bg-primary/15 text-primary font-medium'
                  : 'text-sidebar-foreground hover:bg-sidebar-accent hover:text-sidebar-accent-foreground'
              }`}
            >
              <Icon className="w-4 h-4 shrink-0" />
              {label}
              {isActive && (
                <span className="ml-auto w-1.5 h-1.5 rounded-full bg-primary" />
              )}
            </Link>
          )
        })}
      </nav>

      {/* User + Logout */}
      <div className="border-t border-border px-3 py-4">
        {user && (
          <div className="flex items-center gap-2.5 px-2 mb-3">
            <div className="flex items-center justify-center w-7 h-7 rounded-full bg-secondary">
              <User className="w-3.5 h-3.5 text-muted-foreground" />
            </div>
            <div className="flex flex-col min-w-0">
              <p className="text-xs font-medium text-sidebar-foreground truncate">{user.email}</p>
              <p className="text-xs text-muted-foreground">{user.role}</p>
            </div>
          </div>
        )}
        <button
          onClick={onLogout}
          className="flex items-center gap-2.5 w-full px-3 py-2 rounded-md text-sm text-muted-foreground hover:bg-destructive/15 hover:text-destructive transition-colors"
        >
          <LogOut className="w-4 h-4 shrink-0" />
          Logout
        </button>
      </div>
    </aside>
  )
}
