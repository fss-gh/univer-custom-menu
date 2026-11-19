import React from 'react';

/** 打印缩放：页面 + 四角向外扩展 */
export function PageLayoutScaleIcon() {
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
            <rect x="4.5" y="4.5" width="7" height="7" rx="1" />
            <path d="M2.5 6V2.5H6M13.5 6V2.5H10M2.5 10v3.5H6M13.5 10v3.5H10" />
        </svg>
    );
}
