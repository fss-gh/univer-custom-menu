import React from 'react';

/** 分页符：纸张 + 横向虚线分隔 */
export function PageLayoutPageBreaksIcon() {
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
            <rect x="2.5" y="1.6" width="11" height="12.8" rx="1.2" />
            <path d="M2.5 8h11" strokeDasharray="2.1 1.5" />
            <path d="M5.4 4.3h2.2M8.4 4.3h2.2M5.4 11.7h2.2M8.4 11.7h2.2" />
        </svg>
    );
}
