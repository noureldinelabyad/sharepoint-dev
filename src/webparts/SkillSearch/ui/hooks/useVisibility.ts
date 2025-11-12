// src/webparts/SkillSearch/ui/hooks/useVisibility.ts
import * as React from 'react';

export function useVisibility<T extends Element>(rootMargin = '300px') {
  const ref = React.useRef<T | null>(null);
  const [visible, setVisible] = React.useState(false);

  React.useEffect(() => {
    const el = ref.current;
    if (!el || visible) return;
    const io = new IntersectionObserver(([entry]) => {
      if (entry.isIntersecting) { setVisible(true); io.disconnect(); }
    }, { rootMargin });
    io.observe(el);
    return () => io.disconnect();
  }, [visible]);

  return { ref, visible } as const;
}
