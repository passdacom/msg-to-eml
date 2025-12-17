/**
 * MSG to EML Converter - Frontend Application
 */

class MSGConverter {
    constructor() {
        this.files = new Map(); // fileId -> { file, status, emlName }
        this.pendingFiles = []; // Files waiting to be uploaded

        this.initElements();
        this.bindEvents();
    }

    initElements() {
        this.uploadZone = document.getElementById('uploadZone');
        this.fileInput = document.getElementById('fileInput');
        this.fileListContainer = document.getElementById('fileListContainer');
        this.fileList = document.getElementById('fileList');
        this.fileCount = document.getElementById('fileCount');
        this.actionButtons = document.getElementById('actionButtons');
        this.convertAllBtn = document.getElementById('convertAllBtn');
        this.downloadAllBtn = document.getElementById('downloadAllBtn');
        this.clearAllBtn = document.getElementById('clearAllBtn');
    }

    bindEvents() {
        // Upload zone click
        this.uploadZone.addEventListener('click', () => this.fileInput.click());

        // File input change
        this.fileInput.addEventListener('change', (e) => this.handleFiles(e.target.files));

        // Drag and drop
        this.uploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.uploadZone.classList.add('drag-over');
        });

        this.uploadZone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            this.uploadZone.classList.remove('drag-over');
        });

        this.uploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            this.uploadZone.classList.remove('drag-over');
            this.handleFiles(e.dataTransfer.files);
        });

        // Action buttons
        this.convertAllBtn.addEventListener('click', () => this.convertAll());
        this.downloadAllBtn.addEventListener('click', () => this.downloadAll());
        this.clearAllBtn.addEventListener('click', () => this.clearAll());
    }

    handleFiles(fileList) {
        const validFiles = Array.from(fileList).filter(file =>
            file.name.toLowerCase().endsWith('.msg')
        );

        if (validFiles.length === 0) {
            this.showToast('MSG 파일만 업로드할 수 있습니다', 'error');
            return;
        }

        validFiles.forEach(file => {
            const tempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
            this.files.set(tempId, {
                file: file,
                name: file.name,
                size: file.size,
                status: 'pending',
                fileId: null,
                emlName: null
            });
        });

        this.updateUI();
        this.fileInput.value = '';
    }

    async convertAll() {
        const pendingFiles = Array.from(this.files.entries())
            .filter(([_, data]) => data.status === 'pending');

        if (pendingFiles.length === 0) {
            this.showToast('변환할 파일이 없습니다', 'info');
            return;
        }

        this.convertAllBtn.disabled = true;

        for (const [tempId, data] of pendingFiles) {
            await this.convertFile(tempId, data);
        }

        this.convertAllBtn.disabled = false;
        this.updateUI();
    }

    async convertFile(tempId, data) {
        // Update status to converting
        data.status = 'converting';
        this.updateFileItem(tempId, data);

        const formData = new FormData();
        formData.append('file', data.file);

        try {
            const response = await fetch('/api/convert', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                data.status = 'success';
                data.fileId = result.file_id;
                data.emlName = result.eml_name;
            } else {
                data.status = 'error';
                data.error = result.error;
            }
        } catch (error) {
            data.status = 'error';
            data.error = '변환 중 오류가 발생했습니다';
        }

        this.updateFileItem(tempId, data);
        this.updateUI();
    }

    async downloadFile(tempId, data) {
        if (data.status !== 'success' || !data.fileId) return;

        const link = document.createElement('a');
        link.href = `/api/download/${data.fileId}`;
        link.download = data.emlName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    async downloadAll() {
        const successFiles = Array.from(this.files.entries())
            .filter(([_, data]) => data.status === 'success' && data.fileId);

        if (successFiles.length === 0) {
            this.showToast('다운로드할 파일이 없습니다', 'info');
            return;
        }

        if (successFiles.length === 1) {
            // 파일이 하나면 개별 다운로드
            this.downloadFile(successFiles[0][0], successFiles[0][1]);
            return;
        }

        // 여러 파일은 ZIP으로
        this.downloadAllBtn.disabled = true;

        try {
            const fileIds = successFiles.map(([_, data]) => data.fileId);

            const response = await fetch('/api/download-all', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ file_ids: fileIds })
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = 'converted_emails.zip';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                window.URL.revokeObjectURL(url);
            }
        } catch (error) {
            this.showToast('다운로드 중 오류가 발생했습니다', 'error');
        }

        this.downloadAllBtn.disabled = false;
    }

    removeFile(tempId) {
        this.files.delete(tempId);
        this.updateUI();
    }

    async clearAll() {
        const fileIds = Array.from(this.files.values())
            .filter(data => data.fileId)
            .map(data => data.fileId);

        if (fileIds.length > 0) {
            try {
                await fetch('/api/clear', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ file_ids: fileIds })
                });
            } catch (error) {
                // Ignore cleanup errors
            }
        }

        this.files.clear();
        this.updateUI();
    }

    updateUI() {
        const count = this.files.size;
        this.fileCount.textContent = count;

        if (count > 0) {
            this.fileListContainer.classList.add('visible');
            this.actionButtons.classList.add('visible');
        } else {
            this.fileListContainer.classList.remove('visible');
            this.actionButtons.classList.remove('visible');
        }

        // Render file list
        this.fileList.innerHTML = '';
        this.files.forEach((data, tempId) => {
            this.fileList.appendChild(this.createFileItem(tempId, data));
        });

        // Update download button state
        const hasSuccess = Array.from(this.files.values()).some(d => d.status === 'success');
        this.downloadAllBtn.disabled = !hasSuccess;

        // Update convert button text
        const pendingCount = Array.from(this.files.values()).filter(d => d.status === 'pending').length;
        if (pendingCount > 0) {
            this.convertAllBtn.innerHTML = `
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polyline points="23 4 23 10 17 10"/>
                    <polyline points="1 20 1 14 7 14"/>
                    <path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/>
                </svg>
                ${pendingCount}개 변환
            `;
            this.convertAllBtn.disabled = false;
        } else {
            this.convertAllBtn.innerHTML = `
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polyline points="20 6 9 17 4 12"/>
                </svg>
                변환 완료
            `;
            this.convertAllBtn.disabled = true;
        }
    }

    createFileItem(tempId, data) {
        const item = document.createElement('div');
        item.className = 'file-item';
        item.dataset.id = tempId;

        const statusHTML = this.getStatusHTML(data);
        const actionsHTML = this.getActionsHTML(tempId, data);

        item.innerHTML = `
            <div class="file-icon">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
                    <polyline points="22,6 12,13 2,6"/>
                </svg>
            </div>
            <div class="file-info">
                <div class="file-name">${this.escapeHTML(data.name)}</div>
                <div class="file-size">${this.formatSize(data.size)}</div>
            </div>
            <div class="file-status">
                ${statusHTML}
                ${actionsHTML}
            </div>
        `;

        // Bind action events
        const downloadBtn = item.querySelector('.btn-download');
        if (downloadBtn) {
            downloadBtn.addEventListener('click', () => this.downloadFile(tempId, data));
        }

        const removeBtn = item.querySelector('.btn-remove');
        if (removeBtn) {
            removeBtn.addEventListener('click', () => this.removeFile(tempId));
        }

        return item;
    }

    updateFileItem(tempId, data) {
        const item = this.fileList.querySelector(`[data-id="${tempId}"]`);
        if (item) {
            const statusContainer = item.querySelector('.file-status');
            statusContainer.innerHTML = `
                ${this.getStatusHTML(data)}
                ${this.getActionsHTML(tempId, data)}
            `;

            // Rebind events
            const downloadBtn = item.querySelector('.btn-download');
            if (downloadBtn) {
                downloadBtn.addEventListener('click', () => this.downloadFile(tempId, data));
            }

            const removeBtn = item.querySelector('.btn-remove');
            if (removeBtn) {
                removeBtn.addEventListener('click', () => this.removeFile(tempId));
            }
        }
    }

    getStatusHTML(data) {
        switch (data.status) {
            case 'pending':
                return `
                    <span class="status-badge pending">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <circle cx="12" cy="12" r="10"/>
                        </svg>
                        대기 중
                    </span>
                `;
            case 'converting':
                return `
                    <span class="status-badge converting">
                        <span class="spinner"></span>
                        변환 중
                    </span>
                `;
            case 'success':
                return `
                    <span class="status-badge success">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <polyline points="20 6 9 17 4 12"/>
                        </svg>
                        완료
                    </span>
                `;
            case 'error':
                return `
                    <span class="status-badge error" title="${this.escapeHTML(data.error || '')}">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <circle cx="12" cy="12" r="10"/>
                            <line x1="15" y1="9" x2="9" y2="15"/>
                            <line x1="9" y1="9" x2="15" y2="15"/>
                        </svg>
                        오류
                    </span>
                `;
            default:
                return '';
        }
    }

    getActionsHTML(tempId, data) {
        let html = '';

        if (data.status === 'success') {
            html += `
                <button class="btn btn-icon btn-ghost btn-download" title="다운로드">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/>
                        <line x1="12" y1="15" x2="12" y2="3"/>
                    </svg>
                </button>
            `;
        }

        html += `
            <button class="btn btn-icon btn-ghost btn-remove" title="삭제">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        `;

        return html;
    }

    formatSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }

    escapeHTML(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    showToast(message, type = 'info') {
        // Simple toast notification
        const toast = document.createElement('div');
        toast.style.cssText = `
            position: fixed;
            bottom: 20px;
            left: 50%;
            transform: translateX(-50%);
            padding: 12px 24px;
            background: ${type === 'error' ? 'rgba(239, 68, 68, 0.9)' : 'rgba(139, 92, 246, 0.9)'};
            color: white;
            border-radius: 8px;
            font-size: 14px;
            z-index: 1000;
            animation: fadeIn 0.3s ease;
        `;
        toast.textContent = message;
        document.body.appendChild(toast);

        setTimeout(() => {
            toast.style.animation = 'fadeOut 0.3s ease';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }
}

// Add toast animations
const style = document.createElement('style');
style.textContent = `
    @keyframes fadeIn {
        from { opacity: 0; transform: translateX(-50%) translateY(20px); }
        to { opacity: 1; transform: translateX(-50%) translateY(0); }
    }
    @keyframes fadeOut {
        from { opacity: 1; transform: translateX(-50%) translateY(0); }
        to { opacity: 0; transform: translateX(-50%) translateY(20px); }
    }
`;
document.head.appendChild(style);

// Initialize app
document.addEventListener('DOMContentLoaded', () => {
    window.converter = new MSGConverter();
});
