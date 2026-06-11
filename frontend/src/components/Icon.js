import React from "react";

const icons = {
  shield: (
    <path d="M12 3l7 3v5c0 4.3-2.7 7.9-7 10-4.3-2.1-7-5.7-7-10V6l7-3z" />
  ),
  tool: (
    <>
      <path d="M14.5 5.5a4.5 4.5 0 0 0 4 6.4l-6.6 6.6a2.1 2.1 0 0 1-3 0l-3.4-3.4a2.1 2.1 0 0 1 0-3l6.6-6.6a4.5 4.5 0 0 0 2.4 0z" />
      <path d="M6 18l-2 2" />
    </>
  ),
  globe: (
    <>
      <circle cx="12" cy="12" r="9" />
      <path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18" />
    </>
  ),
  refresh: (
    <>
      <path d="M20 12a8 8 0 0 1-13.7 5.6" />
      <path d="M4 12A8 8 0 0 1 17.7 6.4" />
      <path d="M17 2v5h-5M7 22v-5h5" />
    </>
  ),
  home: (
    <>
      <path d="M4 10.5L12 4l8 6.5" />
      <path d="M6.5 9.5V20h11V9.5" />
      <path d="M10 20v-6h4v6" />
    </>
  ),
  library: (
    <>
      <path d="M5 5h4v14H5zM10 5h4v14h-4z" />
      <path d="M15 6l3.5 1 1.5 11-4 1-1-13z" />
    </>
  ),
  folder: (
    <path d="M3.5 6.5h6l2 2h9v9a2 2 0 0 1-2 2h-15v-13z" />
  ),
  file: (
    <>
      <path d="M7 3.5h7l4 4V20H7z" />
      <path d="M14 3.5v4h4" />
    </>
  ),
  doc: (
    <>
      <path d="M7 3.5h7l4 4V20H7z" />
      <path d="M14 3.5v4h4M9.5 12h5M9.5 15h5" />
    </>
  ),
  sheet: (
    <>
      <path d="M7 3.5h7l4 4V20H7z" />
      <path d="M14 3.5v4h4M9 12h6M9 15h6M11 10v7" />
    </>
  ),
  slides: (
    <>
      <path d="M5 5h14v10H5z" />
      <path d="M9 19h6M12 15v4" />
    </>
  ),
  image: (
    <>
      <path d="M4.5 5.5h15v13h-15z" />
      <path d="M8 15l3-3 2 2 2.5-3 3.5 4" />
      <circle cx="9" cy="9" r="1" />
    </>
  ),
  paperclip: (
    <path d="M8 12.5l5.6-5.6a3 3 0 0 1 4.2 4.2l-7 7a4.5 4.5 0 0 1-6.4-6.4l7.2-7.2" />
  ),
  warning: (
    <>
      <path d="M12 4l9 16H3z" />
      <path d="M12 9v5M12 17h.01" />
    </>
  ),
  check: (
    <path d="M5 12.5l4 4L19 6.5" />
  ),
};

function Icon({ name, className = "", size = 18, title }) {
  return (
    <svg
      aria-hidden={title ? undefined : "true"}
      aria-label={title}
      className={`app-icon ${className}`.trim()}
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      focusable="false"
    >
      {icons[name] || icons.file}
    </svg>
  );
}

export default Icon;
