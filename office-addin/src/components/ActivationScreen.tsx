import React, { useState } from 'react';

interface ActivationScreenProps {
  onActivate: (key: string) => void;
  backendUrl: string;
}

export function ActivationScreen({ onActivate, backendUrl }: ActivationScreenProps) {
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
      alignItems: 'center',
      justifyContent: 'center',
      height: '100vh',
      padding: '20px',
      textAlign: 'center',
      backgroundColor: '#f9f9f9',
      fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif'
    }}>
      <h2 style={{ color: '#2c3e50', marginBottom: '10px' }}>تنشيط الإضافة</h2>
      <p style={{ color: '#7f8c8d', marginBottom: '20px', fontSize: '14px' }}>
        يرجى إدخال مفتاح التنشيط الخاص بك للوصول إلى خدمات الذكاء الاصطناعي.
      </p>
      
      <input
        type="text"
        value={keyInput}
        onChange={(e) => setKeyInput(e.target.value)}
        placeholder="أدخل مفتاح التنشيط (مثال: RV-...)"
        style={{
          width: '100%',
          maxWidth: '300px',
          padding: '12px',
          borderRadius: '6px',
          border: '1px solid #bdc3c7',
          marginBottom: '15px',
          fontSize: '16px',
          textAlign: 'center'
        }}
        disabled={isVerifying}
      />
      
      {error && (
        <div style={{ color: '#e74c3c', marginBottom: '15px', fontSize: '13px' }}>
          {error}
        </div>
      )}
      
      <button
        onClick={handleVerify}
        disabled={isVerifying}
        style={{
          backgroundColor: '#3498db',
          color: 'white',
          border: 'none',
          padding: '12px 24px',
          borderRadius: '6px',
          cursor: isVerifying ? 'not-allowed' : 'pointer',
          fontSize: '16px',
          fontWeight: 'bold',
          opacity: isVerifying ? 0.7 : 1,
          width: '100%',
          maxWidth: '300px'
        }}
      >
        {isVerifying ? 'جاري التحقق...' : 'تنشيط'}
      </button>

      <div style={{ marginTop: '30px', fontSize: '12px', color: '#95a5a6' }}>
        للحصول على مفتاح تنشيط، يرجى التواصل مع الدعم الفني.
      </div>
    </div>
  );
}
