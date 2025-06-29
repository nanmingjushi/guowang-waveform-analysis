new Vue({
    el: '#app',
    data: {
        selectedFile: null,
        selectedTemplateFile: null,
        selectedImages: [],
        isDragging: false,
        isLoading: false,
        error: null,
        success: null,
        downloadUrl: null,
        // 其它表单字段
        fields: {
            reportNo: '',
            client: '',
            addressOfClient: '',
            applicant: '',
            addressOfApplicant: '',
            testSite: '',
            voltage: '',
            spot: '',
            startYear: '', startMonth: '', startDay: '', startHour: '', startMinute: '',
            endYear: '', endMonth: '', endDay: '', endHour: '', endMinute: '',
            // 仪器信息支持多行
            measurements: [
                { measurement: '', certificateNo: '', certificateDate: '' }
            ]
        },
        // 仪器选项，value为仪器名
        measurementDict: {
            "FLUKE-1777": { certificateNo: "11111", certificateDate: "2025.6.23" },
            "选项2": { certificateNo: "22222", certificateDate: "2026.1.1" },
            "选项3": { certificateNo: "33333", certificateDate: "2027.3.15" }
        }
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
        handleImagesChange(e) {
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
        // measurement选项切换时自动填充编号与有效期
        onMeasurementChange(idx) {
            const item = this.fields.measurements[idx];
            const dict = this.measurementDict[item.measurement];
            if (dict) {
                this.$set(item, 'certificateNo', dict.certificateNo);
                this.$set(item, 'certificateDate', dict.certificateDate);
            } else {
                this.$set(item, 'certificateNo', '');
                this.$set(item, 'certificateDate', '');
            }
        },
        addMeasurement() {
            this.fields.measurements.push({ measurement: '', certificateNo: '', certificateDate: '' });
        },
        removeMeasurement(idx) {
            this.fields.measurements.splice(idx, 1);
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
            // 图片上传
            this.selectedImages.forEach(file => {
                formData.append('images', file, file.name);
            });
            // 其它字段添加
            Object.keys(this.fields).forEach(key => {
                if (key === 'measurements') {
                    // 多行仪器：序列化为JSON字符串
                    formData.append('measurements', JSON.stringify(this.fields.measurements));
                } else {
                    formData.append(key, this.fields[key]);
                }
            });

            axios.post('/upload', formData, {
                headers: { 'Content-Type': 'multipart/form-data' }
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
