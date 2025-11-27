import { useEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent } from 'react'
import './App.css'

type TeamRecord = {
  teamid: string
  name: string
  bu?: string
  email?: string
  description?: string
  [key: string]: unknown
}

type RoleRecord = {
  roleid: string
  name: string
  teamid: string
  privilege?: string
  environment?: string
  [key: string]: unknown
}

type TeamsResponse = {
  value: TeamRecord[]
}

type TabularRow = {
  teamName: string
  teamId: string
  bu: string
  email: string
  roleName: string
  privilege: string
  environment: string
}

const mockTeamsResponse: TeamsResponse = {
  value: [
    {
      teamid: '5d0a8577-4d4a-4db9-8a41-10a10d998a6a',
      name: 'APAC Sales Squad',
      bu: 'Sales',
      email: 'apac-sales@dataverse.demo',
      description: 'APAC向け商談と案件引継ぎを担当',
    },
    {
      teamid: '70608fd8-7226-4b5f-82c9-f252fbb5b2c4',
      name: 'Customer Care Core',
      bu: 'Service',
      email: 'care-core@dataverse.demo',
      description: '一次受電とCS調査を担当するサポートチーム',
    },
    {
      teamid: 'b1c8f265-12d0-42a3-8a48-1d679bc08159',
      name: 'DataOps Guild',
      bu: 'Operations',
      email: 'dataops@dataverse.demo',
      description: 'Dataverse連携とCleansingを担う分析チーム',
    },
  ],
}

const mockRoleResponses: Record<string, RoleRecord[]> = {
  '5d0a8577-4d4a-4db9-8a41-10a10d998a6a': [
    {
      roleid: '0bb9e9d8-6051-4a34-96d4-0605c7221599',
      name: 'Sales Manager',
      teamid: '5d0a8577-4d4a-4db9-8a41-10a10d998a6a',
      privilege: 'Read/Write + Share',
      environment: 'PROD',
    },
    {
      roleid: 'df93d5c7-bccb-4638-b6f4-d2a7c4ede94d',
      name: 'Quote Approver',
      teamid: '5d0a8577-4d4a-4db9-8a41-10a10d998a6a',
      privilege: 'Approve Quotes',
      environment: 'PROD',
    },
  ],
  '70608fd8-7226-4b5f-82c9-f252fbb5b2c4': [
    {
      roleid: 'b3f1a1e9-3588-4b73-864c-e0c668d75fa0',
      name: 'Case Agent',
      teamid: '70608fd8-7226-4b5f-82c9-f252fbb5b2c4',
      privilege: 'Create/Update Cases',
      environment: 'PROD',
    },
  ],
  'b1c8f265-12d0-42a3-8a48-1d679bc08159': [
    {
      roleid: '3f6c64fc-7189-47d7-8f04-6bd9dfe8cc1a',
      name: 'System Customizer',
      teamid: 'b1c8f265-12d0-42a3-8a48-1d679bc08159',
      privilege: 'Solution Export',
      environment: 'SANDBOX',
    },
    {
      roleid: '5e3b0623-5436-4fd4-8efa-0ff2f634eb07',
      name: 'Environment Maker',
      teamid: 'b1c8f265-12d0-42a3-8a48-1d679bc08159',
      privilege: 'Maker + PowerFX',
      environment: 'SANDBOX',
    },
  ],
}

const defaultTeamsEndpoint =
  'https://your-org.crm.dynamics.com/api/data/v9.2/teams?$select=teamid,name,businessunitid'
const defaultRolesEndpointTemplate =
  'https://your-org.crm.dynamics.com/api/data/v9.2/teams/{teamId}/teammembership_association?$select=roleid,name'

const buildHeaders = (token?: string): HeadersInit => {
  const headers: HeadersInit = {
    Accept: 'application/json',
    'OData-Version': '4.0',
  }
  if (token) {
    headers.Authorization = `Bearer ${token}`
  }
  return headers
}

const resolveRoleUrl = (template: string, teamId: string) => {
  if (template.includes('{{teamId}}')) {
    return template.replaceAll('{{teamId}}', teamId)
  }
  if (template.includes('{teamId}')) {
    return template.replaceAll('{teamId}', teamId)
  }
  const connector = template.includes('?') ? '&' : '?'
  return `${template}${connector}teamid=${encodeURIComponent(teamId)}`
}

function App() {
  const [teamsEndpoint, setTeamsEndpoint] = useState(defaultTeamsEndpoint)
  const [rolesEndpointTemplate, setRolesEndpointTemplate] = useState(
    defaultRolesEndpointTemplate,
  )
  const [token, setToken] = useState('')
  const [teams, setTeams] = useState<TeamRecord[]>([])
  const [rolesByTeam, setRolesByTeam] = useState<Record<string, RoleRecord[]>>(
    {},
  )
  const [teamsLoading, setTeamsLoading] = useState(false)
  const [rolesLoading, setRolesLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [lastTeamsSource, setLastTeamsSource] = useState<
    'mock' | 'api' | 'json' | null
  >(null)
  const [lastRolesSource, setLastRolesSource] = useState<'mock' | 'api' | null>(
    null,
  )
  const [copyMessage, setCopyMessage] = useState<string | null>(null)
  const copyTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  const combinedRows = useMemo(
    () =>
      teams.map((team) => ({
        team,
        roles: rolesByTeam[team.teamid] ?? [],
      })),
    [teams, rolesByTeam],
  )

  const tabularRows = useMemo<TabularRow[]>(() => {
    if (combinedRows.length === 0) {
      return []
    }
    return combinedRows.flatMap(({ team, roles }) => {
      if (roles.length === 0) {
        return [
          {
            teamName: team.name ?? '',
            teamId: team.teamid ?? '',
            bu: team.bu ?? '',
            email: team.email ?? '',
            roleName: '',
            privilege: '',
            environment: '',
          },
        ]
      }
      return roles.map((role) => ({
        teamName: team.name ?? '',
        teamId: team.teamid ?? '',
        bu: team.bu ?? '',
        email: team.email ?? '',
        roleName: role.name ?? '',
        privilege: role.privilege ?? '',
        environment: role.environment ?? '',
      }))
    })
  }, [combinedRows])

  const loadMockData = () => {
    setError(null)
    setTeams(mockTeamsResponse.value)
    setRolesByTeam(mockRoleResponses)
    setLastTeamsSource('mock')
    setLastRolesSource('mock')
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
      setRolesByTeam({})
      setLastTeamsSource('json')
      setLastRolesSource(null)
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
      'Email',
      'Role Name',
      'Privilege',
      'Environment',
    ]
      .join('\t')
      .trim()

    const lines = tabularRows.map((row) =>
      [
        row.teamName,
        row.teamId,
        row.bu,
        row.email,
        row.roleName,
        row.privilege,
        row.environment,
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

  const fetchTeams = async () => {
    setTeamsLoading(true)
    setError(null)
    setLastTeamsSource(null)
    try {
      const teamsResponse = await fetch(teamsEndpoint, {
        headers: buildHeaders(token),
      })
      if (!teamsResponse.ok) {
        throw new Error(
          `Teams APIエラー: ${teamsResponse.status} ${teamsResponse.statusText}`,
        )
      }
      const teamsJson = (await teamsResponse.json()) as TeamsResponse
      const teamList = teamsJson.value ?? []
      setTeams(teamList)
      setRolesByTeam({})
      setLastTeamsSource('api')
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err))
    } finally {
      setTeamsLoading(false)
    }
  }

  const fetchRoles = async () => {
    if (teams.length === 0) {
      setError('先にTeams APIを実行してチーム一覧を取得してください。')
      return
    }
    setRolesLoading(true)
    setError(null)
    setLastRolesSource(null)
    try {
      const roleMap: Record<string, RoleRecord[]> = {}
      for (const team of teams) {
        const roleUrl = resolveRoleUrl(rolesEndpointTemplate, team.teamid)
        const roleResponse = await fetch(roleUrl, {
          headers: buildHeaders(token),
        })
        if (!roleResponse.ok) {
          roleMap[team.teamid] = []
          continue
        }
        const roleJson = (await roleResponse.json()) as {
          value?: RoleRecord[]
          role?: RoleRecord[]
        }
        roleMap[team.teamid] = roleJson.value ?? roleJson.role ?? []
      }
      setRolesByTeam(roleMap)
      setLastRolesSource('api')
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err))
    } finally {
      setRolesLoading(false)
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
          Teams APIで取得した一覧を保持し、teamidを使ってRoles APIを繰り返し実行。
          シンプルな表形式でExcelへコピーできます。
        </p>
      </div>

      <section className="card">
        <h2>1. API設定と実行</h2>
        <p className="caption">
          <code>{'{teamId}'}</code> または <code>{'{{teamId}}'}</code> をURLに含めるとチームIDが差し込まれます。
        </p>
        <div className="controls-grid">
          <label>
            Teams API URL
            <input
              type="text"
              value={teamsEndpoint}
              onChange={(event) => setTeamsEndpoint(event.target.value)}
              placeholder="https://org.crm.dynamics.com/api/data/v9.2/teams?... "
            />
          </label>
          <label>
            Roles API URLテンプレート
            <input
              type="text"
              value={rolesEndpointTemplate}
              onChange={(event) => setRolesEndpointTemplate(event.target.value)}
              placeholder="https://org.crm.dynamics.com/api/.../{teamId}"
            />
          </label>
          <label>
            Bearer Token (任意)
            <input
              type="password"
              value={token}
              onChange={(event) => setToken(event.target.value)}
              placeholder="Dataverse OAuthトークン"
            />
          </label>
          <label>
            JSONファイル読込
            <input type="file" accept="application/json" onChange={handleUpload} />
          </label>
        </div>
        <div className="actions">
          <button type="button" className="outline" onClick={loadMockData}>
            モックデータを読み込む
          </button>
          <button
            type="button"
            className="primary"
            onClick={fetchTeams}
            disabled={teamsLoading}
          >
            {teamsLoading ? 'Teams取得中...' : 'Teams APIを実行'}
          </button>
          <button
            type="button"
            className="secondary"
            onClick={fetchRoles}
            disabled={rolesLoading || teams.length === 0}
          >
            {rolesLoading ? 'Roles取得中...' : 'Roles APIを実行'}
          </button>
        </div>
        {error && <p className="error">⚠️ {error}</p>}
        <div className="status-stack">
          {lastTeamsSource && (
            <p className="status">
              Teams: {lastTeamsSource === 'mock' ? 'モックデータ' : 'API呼び出し'}
              ／{teams.length}件
            </p>
          )}
          {lastRolesSource && (
            <p className="status">
              Roles: {lastRolesSource === 'mock' ? 'モックデータ' : 'API呼び出し'}
            </p>
          )}
        </div>
      </section>

      <section className="card">
        <div className="table-header">
          <div>
            <div className="heading-row">
              <h2>2. 結果表（Excel貼り付け用）</h2>
              <span className="badge">Teams: {teams.length}</span>
            </div>
            <p className="caption">
              列幅を固定せず、シンプルなタブ区切り形式でコピーできます。
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
                <th>Email</th>
                <th>Role Name</th>
                <th>Privilege</th>
                <th>Environment</th>
              </tr>
            </thead>
            <tbody>
              {tabularRows.length === 0 ? (
                <tr>
                  <td colSpan={7} className="empty-row">
                    まだデータがありません。モックデータを読み込むかAPIを実行してください。
                  </td>
                </tr>
              ) : (
                tabularRows.map((row, index) => (
                  <tr key={`${row.teamId}-${row.roleName}-${index}`}>
                    <td>{row.teamName || '-'}</td>
                    <td>{row.teamId || '-'}</td>
                    <td>{row.bu || '-'}</td>
                    <td>{row.email || '-'}</td>
                    <td>{row.roleName || '-'}</td>
                    <td>{row.privilege || '-'}</td>
                    <td>{row.environment || '-'}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>

      <section className="card">
        <div className="table-header">
          <h2>3. Teams JSONプレビュー</h2>
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
