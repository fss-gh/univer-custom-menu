import React from 'react';

/** 打印区域：打印机 + 虚线打印页 */
export function PageLayoutPrintAreaIcon() {
    return (
        <svg
            width="100%"
            height="100%"
            viewBox="0 0 16 16"
            xmlns="http://www.w3.org/2000/svg"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.3"
            strokeLinecap="round"
            strokeLinejoin="round"
            aria-hidden="true"
            focusable="false"
        >
            <path d="M5 1.8h6v3.4H5z" />
            <path d="M2.5 5.2h11v4.6h-2.5V8.2H5v1.6H2.5z" />
            <path d="M5 10.2h6v4H5z" strokeDasharray="1.9 1.3" />
            <circle cx="11.4" cy="7" r="0.55" fill="currentColor" stroke="none" />
        </svg>
    );
}
