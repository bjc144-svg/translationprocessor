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

// Loading indicator with progress steps
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
            background: rgba(0,0,0,0.7);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 9999;
        ">
            <div style="
                background: white;
                padding: 2.5rem;
                border-radius: 12px;
                text-align: center;
                min-width: 400px;
                box-shadow: 0 10px 40px rgba(0,0,0,0.3);
            ">
                <div style="
                    border: 4px solid #f3f3f3;
                    border-top: 4px solid #61912B;
                    border-radius: 50%;
                    width: 60px;
                    height: 60px;
                    animation: spin 1s linear infinite;
                    margin: 0 auto 1.5rem;
                "></div>
                <h3 style="color: #61912B; margin-bottom: 1rem; font-size: 1.3rem;">Processing Translation</h3>
                <p id="loading-status" style="color: #2d3748; font-weight: 500; margin-bottom: 1.5rem;">${message}</p>

                <div style="background: #f0f0f0; border-radius: 8px; padding: 1rem; text-align: left;">
                    <div class="progress-step" data-step="1">
                        <span style="color: #61912B;">✓</span> Reading document...
                    </div>
                    <div class="progress-step" data-step="2">
                        <span style="color: #ccc;">○</span> Adding headers and footers...
                    </div>
                    <div class="progress-step" data-step="3">
                        <span style="color: #ccc;">○</span> Creating certificates...
                    </div>
                    <div class="progress-step" data-step="4">
                        <span style="color: #ccc;">○</span> Converting to PDF...
                    </div>
                </div>

                <p style="color: #666; font-size: 0.9rem; margin-top: 1rem;">This may take a moment...</p>
            </div>
        </div>
    `;
    document.body.appendChild(loader);

    // Add spin animation and styles
    const style = document.createElement('style');
    style.textContent = `
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .progress-step {
            padding: 0.5rem 0;
            font-size: 0.95rem;
            color: #555;
        }
    `;
    document.head.appendChild(style);

    // Simulate progress through steps
    let currentStep = 1;
    const interval = setInterval(() => {
        currentStep++;
        if (currentStep <= 4) {
            updateProgressStep(currentStep);
        } else {
            clearInterval(interval);
        }
    }, 2000);
}

function updateProgressStep(step) {
    const steps = document.querySelectorAll('.progress-step');
    if (steps[step - 1]) {
        const icon = steps[step - 1].querySelector('span');
        if (icon) {
            icon.style.color = '#61912B';
            icon.textContent = '✓';
        }
    }
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
