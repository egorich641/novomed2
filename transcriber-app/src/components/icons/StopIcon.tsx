import type { SVGProps } from 'react';

const StopIcon = (props: SVGProps<SVGSVGElement>) => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={1.5}
    strokeLinecap="round"
    strokeLinejoin="round"
    {...props}
  >
    <rect x={6} y={6} width={12} height={12} rx={2} />
  </svg>
);

export default StopIcon;
