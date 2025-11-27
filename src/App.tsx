import { useEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent } from 'react'
import './App.css'

type RoleAssociation = {
  roleid?: string
  name?: string
}

type TeamRecord = {
  teamid?: string
  name?: string
  _businessunitid_value?: string
  com_team_category?: number
  teamroles_association?: RoleAssociation[]
  [key: string]: unknown
}

type TeamsResponse = {
  value: TeamRecord[]
}

type TabularRow = {
  teamName: string
  teamId: string
  bu: string
  category: string
  roleId: string
  roleName: string
}

const mockTeamsResponse: TeamsResponse = {
  value: [
    {
      teamid: 'd4d7cd62-e35c-4807-83c1-8651724af010',
      name: 'APAC Sales Squad',
      _businessunitid_value: 'BU-SALES',
      com_team_category: 1001,
      teamroles_association: [
        { roleid: 'role-001', name: 'Sales Manager' },
        { roleid: 'role-099', name: 'Quote Approver' },
      ],
    },
    {
      teamid: '3afbfda6-7409-4412-a89f-5dca3a940079',
      name: 'Customer Care Core',
      _businessunitid_value: 'BU-SERVICE',
      com_team_category: 2003,
      teamroles_association: [{ roleid: 'role-201', name: 'Case Agent' }],
    },
    {
      teamid: '7f74218f-0fd4-4f64-9750-0c2709fcc7af',
      name: 'DataOps Guild',
      _businessunitid_value: 'BU-OPS',
      com_team_category: 3100,
      teamroles_association: [
        { roleid: 'role-310', name: 'System Customizer' },
        { roleid: 'role-311', name: 'Environment Maker' },
      ],
    },
  ],
}

const formatCategory = (value?: number) =>
  typeof value === 'number' ? String(value) : ''

function App() {
  const [teams, setTeams] = useState<TeamRecord[]>([])
  const [error, setError] = useState<string | null>(null)
  const [lastSource, setLastSource] = useState<'mock' | 'json' | null>(null)
  const [copyMessage, setCopyMessage] = useState<string | null>(null)
  const copyTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  const tabularRows = useMemo<TabularRow[]>(() => {
    if (teams.length === 0) {
      return []
    }
    return teams.flatMap((team) => {
      const roles = team.teamroles_association ?? []
      if (roles.length === 0) {
        return [
          {
            teamName: team.name ?? '',
            teamId: team.teamid ?? '',
            bu: team._businessunitid_value ?? '',
            category: formatCategory(team.com_team_category),
            roleId: '',
            roleName: '',
          },
        ]
      }
      return roles.map((role) => ({
        teamName: team.name ?? '',
        teamId: team.teamid ?? '',
        bu: team._businessunitid_value ?? '',
        category: formatCategory(team.com_team_category),
        roleId: role.roleid ?? '',
        roleName: role.name ?? '',
      }))
    })
  }, [teams])

  const loadMockData = () => {
    setError(null)
    setTeams(mockTeamsResponse.value)
    setLastSource('mock')
  }

  const clearData = () => {
    setTeams([])
    setLastSource(null)
    setError(null)
  }

  const handleUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }
    try {
      const text = await file.text()
      const parsed = JSON.parse(text) as TeamsResponse
      const list = parsed.value ?? []
      setTeams(list)
      setLastSource('json')
      setError(null)
    } catch (err) {
      setError(
        err instanceof Error
          ? `JSONファイルの読み込みに失敗しました: ${err.message}`
          : 'JSONファイルの読み込みに失敗しました。',
      )
    } finally {
      event.target.value = ''
    }
  }

  const handleCopyTable = async () => {
    if (tabularRows.length === 0) {
      return
    }
    const sanitize = (value: string) =>
      value.replace(/\t/g, ' ').replace(/\r?\n/g, ' ').trim()
    const header = [
      'Team Name',
      'Team ID',
      'Business Unit',
      'Category',
      'Role ID',
      'Role Name',
    ]
      .join('\t')
      .trim()

    const lines = tabularRows.map((row) =>
      [
        row.teamName,
        row.teamId,
        row.bu,
        row.category,
        row.roleId,
        row.roleName,
      ]
        .map((value) => sanitize(value ?? ''))
        .join('\t'),
    )
    const payload = [header, ...lines].join('\n')
    try {
      await navigator.clipboard.writeText(payload)
      setCopyMessage('表データをコピーしました（TSV形式）')
    } catch (err) {
      setCopyMessage(
        err instanceof Error
          ? `コピーに失敗しました: ${err.message}`
          : 'コピーに失敗しました。手動で選択してください。',
      )
    } finally {
      if (copyTimerRef.current) {
        window.clearTimeout(copyTimerRef.current)
      }
      copyTimerRef.current = window.setTimeout(() => {
        setCopyMessage(null)
      }, 2500)
    }
  }

  useEffect(() => {
    return () => {
      if (copyTimerRef.current) {
        window.clearTimeout(copyTimerRef.current)
      }
    }
  }, [])

  return (
    <main className="app">
      <div>
        <h1>Dataverse Teams & Roles</h1>
        <p className="lead">
          DataverseからエクスポートしたJSON（<code>{'{"value":[]}'}</code>形式）を読み込み、
          チームと紐づくロールをExcel向けの表で確認できます。
        </p>
      </div>

      <section className="card">
        <h2>1. JSONを読み込む</h2>
        <p className="caption">
          各チームに
          <code>_businessunitid_value</code>,{' '}
          <code>com_team_category</code>,{' '}
          <code>teamroles_association</code> を含めてください。
        </p>
        <div className="upload-area">
          <label className="file-input">
            <span>JSONファイル</span>
            <input type="file" accept="application/json" onChange={handleUpload} />
          </label>
          <div className="upload-actions">
            <button type="button" className="primary" onClick={loadMockData}>
              サンプルJSONを読み込む
            </button>
            <button
              type="button"
              className="outline"
              onClick={clearData}
              disabled={teams.length === 0}
            >
              クリア
            </button>
          </div>
        </div>
        {error && <p className="error">⚠️ {error}</p>}
        {lastSource && (
          <p className="status">
            {lastSource === 'mock'
              ? 'サンプルデータを表示中'
              : 'JSONファイルを表示中'}
            ／{teams.length}件
          </p>
        )}
      </section>

      <section className="card">
        <div className="table-header">
          <div>
            <div className="heading-row">
              <h2>2. 結果表（Excel貼り付け用）</h2>
              <span className="badge">Teams: {teams.length}</span>
            </div>
            <p className="caption">
              チームとロールの組み合わせをタブ区切りでコピーできます。
            </p>
          </div>
          <button
            type="button"
            className="outline"
            onClick={handleCopyTable}
            disabled={tabularRows.length === 0}
          >
            テーブルをコピー
          </button>
        </div>
        {copyMessage && <p className="copy-message">{copyMessage}</p>}
        <div className="table-scroll">
          <table className="plain-table">
            <thead>
              <tr>
                <th>Team Name</th>
                <th>Team ID</th>
                <th>Business Unit</th>
                <th>Category</th>
                <th>Role ID</th>
                <th>Role Name</th>
              </tr>
            </thead>
            <tbody>
              {tabularRows.length === 0 ? (
                <tr>
                  <td colSpan={6} className="empty-row">
                    まだデータがありません。JSONファイルを読み込むかサンプルを使用してください。
                  </td>
                </tr>
              ) : (
                tabularRows.map((row, index) => (
                  <tr key={`${row.teamId ?? row.teamName}-${row.roleName}-${index}`}>
                    <td>{row.teamName || '-'}</td>
                    <td>{row.teamId || '-'}</td>
                    <td>{row.bu || '-'}</td>
                    <td>{row.category || '-'}</td>
                    <td>{row.roleId || '-'}</td>
                    <td>{row.roleName || '-'}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>

      <section className="card">
        <div className="table-header">
          <h2>3. JSONプレビュー</h2>
          <span className="badge">{teams.length} 件</span>
        </div>
        <pre className="json-block">
          {JSON.stringify(teams, null, 2) || '// まだデータがありません'}
        </pre>
      </section>
    </main>
  )
}

export default App
