import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'AC-34 Matiala — SIR 2026 Hearing Dashboard',
  description: 'Upload the latest ECI NOTICE_REPORT_PART_WISE workbook and instantly refresh the hearing dashboard.',
}

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en"><body>{children}</body></html>
}
