// Translation Processor Web Application - Client-Side JavaScript

// Auto-hide flash messages after 5 seconds
document.addEventListener('DOMContentLoaded', function() {
    const alerts = document.querySelectorAll('.alert');

    alerts.forEach(alert => {
        setTimeout(() => {
            alert.style.opacity = '0';
            alert.style.transform = 'translateY(-20px)';
            setTimeout(() => alert.remove(), 300);
        }, 5000);
    });
});

// Form validation
function validateForm(formId) {
    const form = document.getElementById(formId);
    if (!form) return true;

    const inputs = form.querySelectorAll('[required]');
    let isValid = true;

    inputs.forEach(input => {
        if (!input.value.trim()) {
            isValid = false;
            input.style.borderColor = '#e53e3e';
        } else {
            input.style.borderColor = '#e2e8f0';
        }
    });

    return isValid;
}

// File size validation
function validateFileSize(input, maxSize = 50 * 1024 * 1024) {
    if (input.files && input.files[0]) {
        const fileSize = input.files[0].size;
        if (fileSize > maxSize) {
            alert('File is too large. Maximum size is 50 MB.');
            input.value = '';
            return false;
        }
    }
    return true;
}

// Loading indicator
function showLoading(message = 'Processing...') {
    const loader = document.createElement('div');
    loader.id = 'loading-overlay';
    loader.innerHTML = `
        <div style="
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 9999;
        ">
            <div style="
                background: white;
                padding: 2rem;
                border-radius: 10px;
                text-align: center;
            ">
                <div style="
                    border: 4px solid #f3f3f3;
                    border-top: 4px solid #667eea;
                    border-radius: 50%;
                    width: 50px;
                    height: 50px;
                    animation: spin 1s linear infinite;
                    margin: 0 auto 1rem;
                "></div>
                <p style="color: #2d3748; font-weight: 500;">${message}</p>
            </div>
        </div>
    `;
    document.body.appendChild(loader);

    // Add spin animation
    const style = document.createElement('style');
    style.textContent = '@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }';
    document.head.appendChild(style);
}

function hideLoading() {
    const loader = document.getElementById('loading-overlay');
    if (loader) {
        loader.remove();
    }
}

// Attach loading to process form
document.addEventListener('DOMContentLoaded', function() {
    const processForm = document.getElementById('processForm');
    if (processForm) {
        processForm.addEventListener('submit', function() {
            showLoading('Processing translation document...');
        });
    }
});
