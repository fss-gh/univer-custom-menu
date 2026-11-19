import React from 'react';

/** 纸张：带折角的纸张 + 内容线 */
export function PageLayoutPaperIcon() {
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
            <path d="M2.5 1.6h7.3l3.2 3.2v9.6H2.5z" />
            <path d="M9.8 1.6v3.2h3.2" />
            <path d="M4.7 7.3h5.4M4.7 9.4h5.4M4.7 11.5h3.2" />
        </svg>
    );
}
