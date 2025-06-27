new Vue({
    el: '#app',
    data: {
        selectedFile: null,
        selectedTemplateFile: null,
        selectedImages: [],         // 新增：多图片数组
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
        handleTemplateFileChange(e) {
            this.selectedTemplateFile = e.target.files[0];
            this.validateFile();
        },
        handleImagesChange(e) { // 新增：处理多图片选择
            this.selectedImages = Array.from(e.target.files);
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

            if (!this.selectedFile || !this.selectedTemplateFile) return;

            const validExcelTypes = ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
            const validDocxTypes = ['application/vnd.openxmlformats-officedocument.wordprocessingml.document'];

            const excelFileType = this.selectedFile.type;
            const docxFileType = this.selectedTemplateFile.type;

            if (!validExcelTypes.includes(excelFileType) &&
                !this.selectedFile.name.match(/\.(xls|xlsx)$/)) {
                this.error = '请上传有效的 Excel 文件 (.xls 或 .xlsx)';
                this.selectedFile = null;
                return;
            }

            if (!validDocxTypes.includes(docxFileType) &&
                !this.selectedTemplateFile.name.match(/\.docx$/)) {
                this.error = '请上传有效的 DOCX 模板文件';
                this.selectedTemplateFile = null;
                return;
            }

            if (this.selectedFile.size > 10 * 1024 * 1024 || this.selectedTemplateFile.size > 10 * 1024 * 1024) {
                this.error = '文件大小不能超过 10MB';
                this.selectedFile = null;
                this.selectedTemplateFile = null;
                return;
            }
        },
        uploadFile() {
            if (!this.selectedFile || !this.selectedTemplateFile) return;

            this.isLoading = true;
            this.error = null;
            this.success = null;
            this.downloadUrl = null;

            const formData = new FormData();
            formData.append('file', this.selectedFile);
            formData.append('templateFile', this.selectedTemplateFile);

            // 新增：多图片上传，images 为后端参数名
            this.selectedImages.forEach(file => {
                formData.append('images', file, file.name);
            });

            axios.post('/upload', formData, {
                headers: {
                    'Content-Type': 'multipart/form-data'
                }
            })
                .then(response => {
                    this.success = '报告生成成功！';
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
            let filename = this.downloadUrl.split('/').pop() || 'report.docx';

            fetch(this.downloadUrl)
                .then(resp => {
                    if (!resp.ok) throw new Error('下载失败');
                    return resp.blob();
                })
                .then(blob => {
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = filename;
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
