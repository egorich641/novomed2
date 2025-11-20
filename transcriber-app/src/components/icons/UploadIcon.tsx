import type { SVGProps } from 'react';

const UploadIcon = (props: SVGProps<SVGSVGElement>) => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={1.5}
    strokeLinecap="round"
    strokeLinejoin="round"
    {...props}
  >
    <path d="m6 9 6-6 6 6" />
    <path d="M12 3v12" />
    <path d="M5 21h14" />
  </svg>
);

export default UploadIcon;
