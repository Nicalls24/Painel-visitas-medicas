export const metadata = {
  title: 'Painel Visitas Médicas',
  description: 'Dashboard de visitas hospitalares',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="pt-BR">
      <body>{children}</body>
    </html>
  )
}
