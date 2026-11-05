import * as React from 'react';
import { makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import { ArrowSortUp16Regular, ArrowSortDown16Regular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  button: {
    background: 'none',
    border: 'none',
    padding: 0,
    margin: 0,
    boxSizing: 'border-box',
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXXS,
    cursor: 'pointer',
    userSelect: 'none',
    font: 'inherit',
    color: 'inherit',
    ':focus-visible': {
      outline: `2px solid ${tokens.colorStrokeFocus2}`,
      outlineOffset: '2px',
      borderRadius: tokens.borderRadiusSmall,
    },
    ':hover': {
      color: tokens.colorNeutralForeground1,
    },
  },
});

export interface SortableHeaderProps {
  label: string;
  /** True when the table is currently sorted by this column. */
  active: boolean;
  ascending: boolean;
  onSort: () => void;
  className?: string;
  style?: React.CSSProperties;
}

// A column header whose sort control is a real <button> (focusable, announced as a
// button) inside a <th> that carries aria-sort. The previous <th role="columnheader"
// tabIndex=0 onClick> pattern had no visible focus style and read the ▲▼ glyphs aloud.
export const SortableHeader: React.FC<SortableHeaderProps> = ({ label, active, ascending, onSort, className, style }) => {
  const styles = useStyles();
  return (
    <th
      className={className}
      style={style}
      aria-sort={!active ? 'none' : ascending ? 'ascending' : 'descending'}
    >
      <button type="button" className={mergeClasses(styles.button)} onClick={onSort}>
        {label}
        {active && (ascending
          ? <ArrowSortUp16Regular aria-hidden="true" />
          : <ArrowSortDown16Regular aria-hidden="true" />)}
      </button>
    </th>
  );
};
