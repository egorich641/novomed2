import type { SVGProps } from 'react';

const PencilIcon = (props: SVGProps<SVGSVGElement>) => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={1.5}
    strokeLinecap="round"
    strokeLinejoin="round"
    {...props}
  >
    <path d="m3 21 3-3 12-12a2.1 2.1 0 0 0-3-3L3 15Z" />
    <path d="m15 6 3 3" />
  </svg>
);

export default PencilIcon;
