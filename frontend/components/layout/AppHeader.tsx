'use client';

type Props = {
  appName: string;
  userName?: string;
  onLogout?: () => void;
};

export function AppHeader({ appName, userName, onLogout }: Props) {
  return (
    <header style={headerStyle}>
      <div className="container" style={innerStyle}>
        <h1 style={{ margin: 0, fontSize: 20 }}>{appName}</h1>
        {userName ? (
          <div className="row">
            <span className="badge">{userName}</span>
            <button onClick={onLogout} style={logoutBtn}>ログアウト</button>
          </div>
        ) : null}
      </div>
    </header>
  );
}

const headerStyle: any = {
  background: 'linear-gradient(135deg, #1a237e 0%, #283593 100%)',
  color: '#fff',
  padding: '14px 0',
  borderBottom: '1px solid rgba(255,255,255,0.15)'
};

const innerStyle: any = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between'
};

const logoutBtn: any = {
  background: 'rgba(255,255,255,0.15)',
  color: '#fff',
  border: '1px solid rgba(255,255,255,0.3)',
  borderRadius: 8,
  padding: '6px 10px',
  cursor: 'pointer'
};
