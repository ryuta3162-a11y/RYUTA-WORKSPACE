'use client';

import { useEffect } from 'react';

export default function PwaRegister() {
  useEffect(() => {
    if (!('serviceWorker' in navigator)) return;
    let refreshing = false;
    const onControllerChange = () => {
      if (refreshing) return;
      refreshing = true;
      window.location.reload();
    };
    navigator.serviceWorker.addEventListener('controllerchange', onControllerChange);
    navigator.serviceWorker
      .register('/sw.js?v=23')
      .then((reg) => {
        if (reg?.update) reg.update();
      })
      .catch(() => {});
    return () => {
      navigator.serviceWorker.removeEventListener('controllerchange', onControllerChange);
    };
  }, []);
  return null;
}
