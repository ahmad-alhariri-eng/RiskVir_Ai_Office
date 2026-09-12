import { useState } from 'react';

interface ActivationScreenProps {
  onActivate: (key: string) => void;
  backendUrl: string;
  theme: 'light' | 'dark';
  onToggleTheme: () => void;
}

export function ActivationScreen({ onActivate, backendUrl, theme, onToggleTheme }: ActivationScreenProps) {
  const [keyInput, setKeyInput] = useState('');
  const [isVerifying, setIsVerifying] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleVerify = async () => {
    const trimmedKey = keyInput.trim();
    if (!trimmedKey) {
      setError('يرجى إدخال مفتاح التنشيط');
      return;
    }

    setIsVerifying(true);
    setError(null);

    try {
      const response = await fetch(`${backendUrl}/verify-license`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'License-Key': trimmedKey
        }
      });

      const data = await response.json().catch(() => ({}));

      if (response.ok && data.status === 'valid') {
        onActivate(trimmedKey);
      } else {
        setError(data.reason || 'مفتاح التنشيط غير صالح أو منتهي الصلاحية');
      }
    } catch (err) {
      console.error('License verification failed:', err);
      setError('حدث خطأ أثناء الاتصال بالخادم. يرجى المحاولة لاحقاً.');
    } finally {
      setIsVerifying(false);
    }
  };

  return (
    <div style={{
      display: 'flex',
      flexDirection: 'column',
      height: '100vh',
      backgroundColor: 'var(--bg-primary)',
      color: 'var(--text-primary)',
      fontFamily: '"Inter", system-ui, sans-serif',
      direction: 'rtl'
    }}>
      {/* Header */}
      <header style={{
        display: 'flex',
        justifyContent: 'space-between',
        padding: '16px 24px',
        borderBottom: '1px solid var(--border-color)',
        alignItems: 'center'
      }}>
        <div style={{ fontWeight: 600, fontSize: '18px' }}>RiskVir AI</div>
        <button 
          onClick={onToggleTheme}
          style={{ 
            background: 'transparent', border: 'none', cursor: 'pointer', fontSize: '18px',
            color: 'var(--text-secondary)'
          }}
          title="تغيير الثيم"
        >
          {theme === 'light' ? '🌙' : '☀️'}
        </button>
      </header>

      {/* Main Content */}
      <main style={{
        flex: 1,
        overflowY: 'auto',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        padding: '40px 20px',
        gap: '40px'
      }}>
        
        {/* Activation Card */}
        <div style={{
          width: '100%',
          maxWidth: '440px',
          backgroundColor: 'var(--card-bg)',
          borderRadius: '16px',
          boxShadow: 'var(--card-shadow)',
          padding: '32px',
          border: '1px solid var(--border-color)',
          textAlign: 'center'
        }}>
          <h2 style={{ margin: '0 0 12px 0', fontSize: '24px', fontWeight: 600 }}>تفعيل الإضافة</h2>
          <p style={{ color: 'var(--text-secondary)', margin: '0 0 24px 0', fontSize: '15px', lineHeight: 1.5 }}>
            للبدء في استخدام مساعد الذكاء الاصطناعي، يرجى إدخال مفتاح التفعيل الخاص بك.
          </p>
          
          <input
            type="text"
            value={keyInput}
            onChange={(e) => setKeyInput(e.target.value)}
            placeholder="أدخل مفتاحك هنا (RV-...)"
            style={{
              width: '100%',
              padding: '14px 16px',
              borderRadius: '8px',
              border: '1px solid var(--border-color)',
              backgroundColor: 'var(--input-bg)',
              color: 'var(--text-primary)',
              marginBottom: '16px',
              fontSize: '15px',
              textAlign: 'center',
              outline: 'none',
              transition: 'border-color 0.2s'
            }}
            onFocus={(e) => e.target.style.borderColor = 'var(--accent-color)'}
            onBlur={(e) => e.target.style.borderColor = 'var(--border-color)'}
            disabled={isVerifying}
          />
          
          {error && (
            <div style={{ 
              color: '#d93025', 
              backgroundColor: '#fce8e6', 
              padding: '10px', 
              borderRadius: '6px',
              marginBottom: '16px', 
              fontSize: '13px' 
            }}>
              {error}
            </div>
          )}
          
          <button
            onClick={handleVerify}
            disabled={isVerifying}
            style={{
              backgroundColor: 'var(--accent-color)',
              color: 'white',
              border: 'none',
              padding: '14px 24px',
              borderRadius: '8px',
              cursor: isVerifying ? 'not-allowed' : 'pointer',
              fontSize: '15px',
              fontWeight: 500,
              width: '100%',
              opacity: isVerifying ? 0.7 : 1,
              transition: 'background-color 0.2s'
            }}
            onMouseOver={(e) => !isVerifying && (e.currentTarget.style.backgroundColor = 'var(--accent-hover)')}
            onMouseOut={(e) => !isVerifying && (e.currentTarget.style.backgroundColor = 'var(--accent-color)')}
          >
            {isVerifying ? 'جاري التحقق...' : 'تفعيل'}
          </button>
        </div>

        {/* Pricing Section */}
        <div style={{
          width: '100%',
          maxWidth: '600px',
          textAlign: 'center'
        }}>
          <h3 style={{ fontSize: '20px', marginBottom: '24px', fontWeight: 600 }}>باقات الاشتراك</h3>
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
            gap: '16px'
          }}>
            {/* Plan 1 */}
            <div style={{
              backgroundColor: 'var(--card-bg)',
              border: '1px solid var(--border-color)',
              borderRadius: '12px',
              padding: '24px 16px',
              display: 'flex',
              flexDirection: 'column',
              gap: '12px'
            }}>
              <div style={{ fontSize: '16px', fontWeight: 500 }}>اشتراك شهري</div>
              <div style={{ fontSize: '28px', fontWeight: 700, color: 'var(--accent-color)' }}>$5</div>
              <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>تفعيل لمدة 30 يوم</div>
            </div>

            {/* Plan 2 */}
            <div style={{
              backgroundColor: 'var(--card-bg)',
              border: '2px solid var(--accent-color)',
              borderRadius: '12px',
              padding: '24px 16px',
              display: 'flex',
              flexDirection: 'column',
              gap: '12px',
              position: 'relative'
            }}>
              <div style={{ 
                position: 'absolute', top: '-10px', left: '50%', transform: 'translateX(-50%)',
                background: 'var(--accent-color)', color: 'white', fontSize: '11px', padding: '2px 8px', borderRadius: '10px'
              }}>الأكثر طلباً</div>
              <div style={{ fontSize: '16px', fontWeight: 500 }}>اشتراك 3 أشهر</div>
              <div style={{ fontSize: '28px', fontWeight: 700, color: 'var(--accent-color)' }}>$10</div>
              <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>تفعيل لمدة 90 يوم</div>
            </div>

            {/* Plan 3 */}
            <div style={{
              backgroundColor: 'var(--card-bg)',
              border: '1px solid var(--border-color)',
              borderRadius: '12px',
              padding: '24px 16px',
              display: 'flex',
              flexDirection: 'column',
              gap: '12px'
            }}>
              <div style={{ fontSize: '16px', fontWeight: 500 }}>اشتراك سنوي</div>
              <div style={{ fontSize: '28px', fontWeight: 700, color: 'var(--accent-color)' }}>$25</div>
              <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>تفعيل لمدة 365 يوم</div>
            </div>
          </div>
        </div>

        {/* Support Button */}
        <div style={{ marginTop: 'auto', marginBottom: '20px' }}>
          <a 
            href="https://t.me/Ahmad_CYP" 
            target="_blank" 
            rel="noopener noreferrer"
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              backgroundColor: '#0088cc', /* Telegram Blue */
              color: 'white',
              textDecoration: 'none',
              padding: '12px 24px',
              borderRadius: '24px',
              fontSize: '15px',
              fontWeight: 500,
              boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
              transition: 'transform 0.2s, box-shadow 0.2s'
            }}
            onMouseOver={(e) => { e.currentTarget.style.transform = 'translateY(-2px)'; e.currentTarget.style.boxShadow = '0 4px 8px rgba(0,0,0,0.15)'; }}
            onMouseOut={(e) => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)'; }}
          >
            <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm4.64 6.8c-.15 1.58-.8 5.42-1.13 7.19-.14.75-.42 1-.68 1.03-.58.05-1.02-.38-1.58-.75-.88-.58-1.38-.94-2.23-1.5-.99-.65-.35-1.01.22-1.59.15-.15 2.71-2.48 2.76-2.69.01-.03.01-.14-.07-.19-.08-.05-.19-.02-.27 0-.12.03-1.98 1.25-5.58 3.68-.53.36-1.01.54-1.44.53-.47-.01-1.38-.27-2.06-.49-.83-.27-1.49-.42-1.43-.89.03-.25.38-.51 1.03-.78 4.04-1.76 6.74-2.92 8.09-3.48 3.85-1.6 4.64-1.88 5.17-1.89.11 0 .37.03.54.17.14.12.18.28.2.45-.02.07-.02.13-.02.19z"/>
            </svg>
            تواصل مع الدعم الفني للحصول على المفتاح
          </a>
        </div>

      </main>
    </div>
  );
}
