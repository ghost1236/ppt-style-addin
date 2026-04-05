import { useState, useRef, useEffect } from 'react';
import { HexColorPicker } from 'react-colorful';
import { Input, makeStyles, tokens } from '@fluentui/react-components';

const useStyles = makeStyles({
  wrapper: {
    position: 'relative',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  swatch: {
    width: '28px',
    height: '28px',
    borderRadius: '4px',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    cursor: 'pointer',
    flexShrink: 0,
  },
  popover: {
    position: 'absolute',
    top: '36px',
    left: 0,
    zIndex: 9999,
    background: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: '8px',
    padding: '12px',
    boxShadow: tokens.shadow16,
  },
  hexInput: {
    width: '90px',
  },
});

interface ColorPickerProps {
  color: string;
  onChange: (color: string) => void;
}

export function ColorPicker({ color, onChange }: ColorPickerProps) {
  const styles = useStyles();
  const [open, setOpen] = useState(false);
  const [hexInput, setHexInput] = useState(color);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setHexInput(color);
  }, [color]);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
      }
    }
    if (open) document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [open]);

  function handleHexInput(value: string) {
    setHexInput(value);
    if (/^#[0-9A-Fa-f]{6}$/.test(value)) {
      onChange(value);
    }
  }

  return (
    <div className={styles.wrapper} ref={ref}>
      <div
        className={styles.swatch}
        style={{ backgroundColor: color }}
        onClick={() => setOpen((o) => !o)}
        title="색상 선택"
      />
      <Input
        className={styles.hexInput}
        value={hexInput}
        onChange={(_, d) => handleHexInput(d.value)}
        size="small"
        placeholder="#333333"
      />
      {open && (
        <div className={styles.popover}>
          <HexColorPicker color={color} onChange={(c) => { onChange(c); setHexInput(c); }} />
        </div>
      )}
    </div>
  );
}
