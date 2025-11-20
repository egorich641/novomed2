import type { SVGProps } from 'react';

const CheckIcon = (props: SVGProps<SVGSVGElement>) => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={1.5}
    strokeLinecap="round"
    strokeLinejoin="round"
    {...props}
  >
    <path d="m5 13 4 4L19 7" />
  </svg>
);

export default CheckIcon;
