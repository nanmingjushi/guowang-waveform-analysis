new Vue({
    el: '#app',
    data: {
        selectedFile: null,
        isDragging: false,
        isLoading: false,
        error: null,
        success: null,
        downloadUrl: null
    },
    methods: {
        handleFileChange(e) {
            this.selectedFile = e.target.files[0];
            this.validateFile();
        },
        dragover() {
            this.isDragging = true;
        },
        dragleave() {
            this.isDragging = false;
        },
        drop(e) {
            this.isDragging = false;
            this.selectedFile = e.dataTransfer.files[0];
            this.$refs.file.files = e.dataTransfer.files;
            this.validateFile();
        },
        validateFile() {
            this.error = null;

            if (!this.selectedFile) return;

            const validTypes = ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
            const fileType = this.selectedFile.type;

            if (!validTypes.includes(fileType) &&
                !this.selectedFile.name.match(/\.(xls|xlsx)$/)) {
                this.error = '请上传有效的Excel文件 (.xls 或 .xlsx)';
                this.selectedFile = null;
                return;
            }

            if (this.selectedFile.size > 10 * 1024 * 1024) {
                this.error = '文件大小不能超过10MB';
                this.selectedFile = null;
                return;
            }
        },
        uploadFile() {
            if (!this.selectedFile) return;

            this.isLoading = true;
            this.error = null;
            this.success = null;
            this.downloadUrl = null;

            const formData = new FormData();
            formData.append('file', this.selectedFile);

            // 这里替换为您的实际API端点
            axios.post('/upload', formData, {
                headers: {
                    'Content-Type': 'multipart/form-data'
                }
            })
                .then(response => {
                    this.success = '报告生成成功！';
                    // 假设返回数据中包含下载URL
                    this.downloadUrl = response.data.downloadUrl;
                })
                .catch(error => {
                    console.error('上传错误:', error);
                    this.error = error.response?.data?.message || '文件上传失败，请重试';
                })
                .finally(() => {
                    this.isLoading = false;
                });
        },
        downloadReport() {
            if (!this.downloadUrl) return;
            // 获取文件名（如 /download/xxx.docx）
            let filename = this.downloadUrl.split('/').pop() || 'report.docx';

            fetch(this.downloadUrl)
                .then(resp => {
                    if (!resp.ok) throw new Error('下载失败');
                    return resp.blob();
                })
                .then(blob => {
                    // 创建下载链接对象
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = filename; // 这行让浏览器弹出“另存为”
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                })
                .catch(err => {
                    this.error = '下载失败，请稍后重试';
                    console.error(err);
                });
        }


    }
});