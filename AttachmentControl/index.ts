import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class AttachmentControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private fileInput: HTMLInputElement;
    private previewContainer: HTMLDivElement;
    private notifyOutputChanged: () => void;
    private fileBase64List: { name: string; data: string }[] = [];

    private readonly maxFileSizeMB: number = 5;
    private readonly acceptedFileTypes: string[] = ["image/", "application/pdf", "text/"];

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;

        const uploadedFileData = context.parameters.uploadedFile.raw;
        const overflowFileData = context.parameters.uploadedFileOverflow?.raw;

        this.fileBase64List = [];

        try {
            if (uploadedFileData) {
                this.fileBase64List = JSON.parse(uploadedFileData);
            }
            if (overflowFileData) {
                this.fileBase64List = this.fileBase64List.concat(JSON.parse(overflowFileData));
            }
        } catch (error) {
            console.error("Error parsing uploaded file data:", error);
        }

        this.createUI();
        this.displayPreviews();
    }

    private createUI(): void {
        this.fileInput = document.createElement("input");
        this.fileInput.type = "file";
        this.fileInput.accept = this.acceptedFileTypes.join(",");
        this.fileInput.style.display = "none";
        this.fileInput.multiple = true;
        this.fileInput.onchange = this.handleFileUpload.bind(this);

        const uploadIcon = this.createUploadIcon();
        uploadIcon.onclick = () => this.fileInput.click();

        const dragDropArea = this.createDragDropArea();

        this.previewContainer = document.createElement("div");
        this.previewContainer.style.border = "1px solid #ddd";
        this.previewContainer.style.padding = "10px";
        this.previewContainer.style.marginTop = "10px";
        this.previewContainer.textContent = "No files uploaded.";

        this.container.appendChild(uploadIcon);
        this.container.appendChild(dragDropArea);
        this.container.appendChild(this.fileInput);
        this.container.appendChild(this.previewContainer);
    }

    private createUploadIcon(): HTMLDivElement {
        const uploadIcon = document.createElement("div");
        uploadIcon.textContent = "📤 Upload Files";
        uploadIcon.style.cursor = "pointer";
        uploadIcon.style.padding = "10px";
        uploadIcon.style.border = "1px solid #ccc";
        uploadIcon.style.textAlign = "center";
        uploadIcon.style.marginBottom = "10px";
        uploadIcon.style.backgroundColor = "#f8f8f8";
        uploadIcon.style.borderRadius = "5px";
        uploadIcon.style.fontWeight = "bold";
        return uploadIcon;
    }

    private createDragDropArea(): HTMLDivElement {
        const dragDropArea = document.createElement("div");
        dragDropArea.textContent = "Drag and drop your files here.";
        dragDropArea.style.border = "2px dashed #ccc";
        dragDropArea.style.padding = "20px";
        dragDropArea.style.textAlign = "center";
        dragDropArea.style.marginTop = "10px";
        dragDropArea.style.borderRadius = "5px";
        dragDropArea.style.fontStyle = "italic";
        dragDropArea.ondragover = (event) => {
            event.preventDefault();
            dragDropArea.style.borderColor = "#0078d4";
        };
        dragDropArea.ondragleave = () => {
            dragDropArea.style.borderColor = "#ccc";
        };
        dragDropArea.ondrop = (event) => {
            event.preventDefault();
            dragDropArea.style.borderColor = "#ccc";
            const files = event.dataTransfer?.files;
            if (files) {
                Array.from(files).forEach((file) => this.processFile(file));
            }
        };
        return dragDropArea;
    }

    private handleFileUpload(event: Event): void {
        const files = (event.target as HTMLInputElement).files;
        if (files) {
            Array.from(files).forEach((file) => this.processFile(file));
        }
    }

    private processFile(file: File): void {
        if (!this.validateFile(file)) {
            return;
        }

        const reader = new FileReader();
        reader.onload = () => {
            const fileData = reader.result as string;
            const newFile = { name: file.name, data: fileData };
            
            // First try to add to main field
            const testList = [...this.fileBase64List, newFile];
            const mainFieldJson = JSON.stringify(testList);
            
            if (mainFieldJson.length <= 1000000) {
                this.fileBase64List.push(newFile);
                this.notifyOutputChanged();
                this.displayPreviews();
                return;
            }

            // If main field is full, check if we can split the file
            const currentMainFieldJson = JSON.stringify(this.fileBase64List);
            const remainingSpace = 1000000 - currentMainFieldJson.length;
            
            if (remainingSpace > 0) {
                // Calculate split points and test if both parts can fit
                const splitPoint = Math.floor(remainingSpace / 2); // Divide by 2 to account for JSON structure
                const part1 = fileData.substring(0, splitPoint);
                const part2 = fileData.substring(splitPoint);
                
                // Test if first part can fit in main field
                const testMainField = [...this.fileBase64List, { name: file.name, data: part1 }];
                const testMainFieldJson = JSON.stringify(testMainField);
                
                if (testMainFieldJson.length <= 1000000) {
                    // Test if second part can fit in overflow
                    const currentOverflow = this.getCurrentOverflow();
                    const testOverflow = [...currentOverflow, { name: file.name, data: part2 }];
                    const testOverflowJson = JSON.stringify(testOverflow);
                    
                    if (testOverflowJson.length <= 1000000) {
                        // Both parts can fit, add them
                        this.fileBase64List.push({ name: file.name, data: part1 });
                        this.fileBase64List.push({ name: file.name, data: part2 });
                        this.notifyOutputChanged();
                        this.displayPreviews();
                        return;
                    }
                }
            }

            // If we can't split or fit in overflow, show error
            alert("Maximum storage capacity reached. Cannot upload more files.");
        };
        reader.readAsDataURL(file);
    }

    private getCurrentOverflow(): { name: string; data: string }[] {
        const mainFieldJson = JSON.stringify(this.fileBase64List);
        if (mainFieldJson.length <= 1000000) {
            return [];
        }

        const part1: { name: string; data: string }[] = [];
        const part2: { name: string; data: string }[] = [];
        let currentLength = 0;

        for (const file of this.fileBase64List) {
            const fileJson = JSON.stringify(file);
            const testLength = currentLength + fileJson.length + 1;
            
            if (testLength <= 1000000) {
                part1.push(file);
                currentLength = testLength;
            } else {
                part2.push(file);
            }
        }

        return part2;
    }

    private validateFile(file: File): boolean {
        const maxSizeBytes = this.maxFileSizeMB * 1024 * 1024;
        if (file.size > maxSizeBytes) {
            alert(`File "${file.name}" exceeds the ${this.maxFileSizeMB} MB size limit.`);
            return false;
        }

        if (!this.acceptedFileTypes.some((type) => file.type.startsWith(type))) {
            alert(`Unsupported file type for "${file.name}".`);
            return false;
        }

        return true;
    }

    private createFullScreenPreview(): HTMLDivElement {
        const fullScreenContainer = document.createElement("div");
        fullScreenContainer.style.position = "fixed";
        fullScreenContainer.style.top = "0";
        fullScreenContainer.style.left = "0";
        fullScreenContainer.style.width = "100vw";
        fullScreenContainer.style.height = "100vh";
        fullScreenContainer.style.backgroundColor = "rgba(0, 0, 0, 0.8)";
        fullScreenContainer.style.display = "none";
        fullScreenContainer.style.zIndex = "999999";
        fullScreenContainer.style.overflow = "auto";
        fullScreenContainer.style.cursor = "default";

        const closeButton = document.createElement("button");
        closeButton.innerHTML = '<svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>';
        closeButton.style.position = "fixed";
        closeButton.style.top = "50px";
        closeButton.style.right = "20px";
        closeButton.style.background = "none";
        closeButton.style.border = "none";
        closeButton.style.cursor = "pointer";
        closeButton.style.padding = "8px";
        closeButton.style.color = "#ff4444";
        closeButton.style.transition = "all 0.2s";
        closeButton.style.zIndex = "1000000";
        closeButton.onmouseover = () => {
            closeButton.style.color = "#ff0000";
            closeButton.style.transform = "scale(1.1)";
        };
        closeButton.onmouseout = () => {
            closeButton.style.color = "#ff4444";
            closeButton.style.transform = "scale(1)";
        };
        closeButton.onclick = () => {
            fullScreenContainer.style.display = "none";
        };

        const previewContent = document.createElement("div");
        previewContent.style.width = "100%";
        previewContent.style.height = "100%";
        previewContent.style.display = "flex";
        previewContent.style.justifyContent = "center";
        previewContent.style.alignItems = "center";
        previewContent.style.padding = "20px";
        previewContent.style.boxSizing = "border-box";
        previewContent.style.cursor = "default";

        fullScreenContainer.appendChild(closeButton);
        fullScreenContainer.appendChild(previewContent);

        fullScreenContainer.addEventListener("click", (event) => {
            if (event.target === fullScreenContainer) {
                fullScreenContainer.style.display = "none";
            }
        });

        const escapeHandler = (event: KeyboardEvent) => {
            if (event.key === "Escape") {
                fullScreenContainer.style.display = "none";
            }
        };
        document.addEventListener("keydown", escapeHandler);

        document.body.appendChild(fullScreenContainer);
        return fullScreenContainer;
    }

    private displayPreviews(): void {
        this.previewContainer.innerHTML = "";

        if (this.fileBase64List.length === 0) {
            return;
        }

        const fullScreenContainer = this.createFullScreenPreview();

        // Group files by name to handle split files
        const fileGroups = new Map<string, { name: string; data: string }[]>();
        this.fileBase64List.forEach(file => {
            if (!fileGroups.has(file.name)) {
                fileGroups.set(file.name, []);
            }
            fileGroups.get(file.name)!.push(file);
        });

        fileGroups.forEach((files, fileName) => {
            const fileContainer = document.createElement("div");
            fileContainer.style.marginBottom = "10px";
            fileContainer.style.position = "relative";
            fileContainer.style.padding = "8px";
            fileContainer.style.border = "1px solid #ddd";
            fileContainer.style.borderRadius = "4px";
            fileContainer.style.backgroundColor = "#f9f9f9";

            const fileLabel = document.createElement("p");
            fileLabel.textContent = `File: ${fileName}`;
            fileLabel.style.margin = "0 0 8px 0";
            fileLabel.style.fontWeight = "bold";

            const buttonContainer = document.createElement("div");
            buttonContainer.style.display = "flex";
            buttonContainer.style.gap = "8px";

            const removeButton = document.createElement("button");
            removeButton.innerHTML = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18"></path><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"></path><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"></path></svg>';
            removeButton.style.background = "none";
            removeButton.style.border = "none";
            removeButton.style.cursor = "pointer";
            removeButton.style.padding = "5px";
            removeButton.style.color = "#666";
            removeButton.style.transition = "all 0.2s";
            removeButton.style.marginLeft = "10px";
            removeButton.style.verticalAlign = "middle";
            removeButton.onmouseover = () => {
                removeButton.style.color = "#ff4444";
                removeButton.style.transform = "scale(1.1)";
            };
            removeButton.onmouseout = () => {
                removeButton.style.color = "#666";
                removeButton.style.transform = "scale(1)";
            };
            removeButton.onclick = () => {
                // Remove all parts of the file
                this.fileBase64List = this.fileBase64List.filter(f => f.name !== fileName);
                this.notifyOutputChanged();
                this.displayPreviews();
            };

            const fullScreenButton = document.createElement("button");
            fullScreenButton.innerHTML = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3"></path></svg>';
            fullScreenButton.style.position = "absolute";
            fullScreenButton.style.top = "5px";
            fullScreenButton.style.right = "5px";
            fullScreenButton.style.background = "none";
            fullScreenButton.style.border = "none";
            fullScreenButton.style.cursor = "pointer";
            fullScreenButton.style.padding = "5px";
            fullScreenButton.style.color = "#666";
            fullScreenButton.style.transition = "color 0.2s";
            fullScreenButton.onmouseover = () => {
                fullScreenButton.style.color = "#0078d4";
            };
            fullScreenButton.onmouseout = () => {
                fullScreenButton.style.color = "#666";
            };
            fullScreenButton.onclick = () => {
                const previewContent = fullScreenContainer.querySelector("div");
                if (previewContent) {
                    previewContent.innerHTML = "";
                    // Combine all parts of the file for preview
                    const combinedData = files.map(f => f.data).join("");
                    if (combinedData.startsWith("data:image/")) {
                        const img = document.createElement("img");
                        img.src = combinedData;
                        img.style.maxWidth = "100%";
                        previewContent.appendChild(img);
                    } else if (combinedData.startsWith("data:application/pdf")) {
                        const iframe = document.createElement("iframe");
                        iframe.src = combinedData;
                        iframe.style.width = "100%";
                        iframe.style.height = "100%";
                        previewContent.appendChild(iframe);
                    } else if (combinedData.startsWith("data:text/")) {
                        const textArea = document.createElement("textarea");
                        textArea.value = atob(combinedData.split(",")[1]);
                        textArea.readOnly = true;
                        textArea.style.width = "100%";
                        textArea.style.height = "100%";
                        previewContent.appendChild(textArea);
                    }
                    fullScreenContainer.style.display = "block";
                }
            };

            buttonContainer.appendChild(removeButton);
            buttonContainer.appendChild(fullScreenButton);
            fileContainer.appendChild(fileLabel);
            fileContainer.appendChild(buttonContainer);

            // Show preview for all file types
            const combinedData = files.map(f => f.data).join("");
            if (combinedData.startsWith("data:image/")) {
                const img = document.createElement("img");
                img.src = combinedData;
                img.style.maxWidth = "100%";
                img.style.maxHeight = "200px";
                img.style.objectFit = "contain";
                fileContainer.appendChild(img);
            } else if (combinedData.startsWith("data:application/pdf")) {
                const iframe = document.createElement("iframe");
                iframe.src = combinedData;
                iframe.style.width = "100%";
                iframe.style.height = "200px";
                iframe.style.border = "none";
                fileContainer.appendChild(iframe);
            } else if (combinedData.startsWith("data:text/")) {
                const textArea = document.createElement("textarea");
                textArea.value = atob(combinedData.split(",")[1]);
                textArea.readOnly = true;
                textArea.style.width = "100%";
                textArea.style.height = "100px";
                textArea.style.resize = "none";
                textArea.style.padding = "8px";
                textArea.style.border = "1px solid #ddd";
                textArea.style.borderRadius = "4px";
                fileContainer.appendChild(textArea);
            }

            this.previewContainer.appendChild(fileContainer);
        });
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const uploadedFileData = context.parameters.uploadedFile.raw;
        const overflowFileData = context.parameters.uploadedFileOverflow?.raw;

        this.fileBase64List = [];
        try {
            if (uploadedFileData) this.fileBase64List = JSON.parse(uploadedFileData);
            if (overflowFileData) this.fileBase64List = this.fileBase64List.concat(JSON.parse(overflowFileData));
        } catch (e) {
            this.fileBase64List = [];
            console.error("Failed to parse file list on updateView");
        }

        this.displayPreviews();
    }

    public getOutputs(): IOutputs {
        // First, try to merge any split files that can fit in the main field
        const mergedFiles: { name: string; data: string }[] = [];
        const remainingFiles: { name: string; data: string }[] = [];
        const fileGroups = new Map<string, { name: string; data: string }[]>();

        // Group files by name to identify split files
        this.fileBase64List.forEach(file => {
            if (!fileGroups.has(file.name)) {
                fileGroups.set(file.name, []);
            }
            fileGroups.get(file.name)!.push(file);
        });

        // Try to merge split files
        fileGroups.forEach((files, fileName) => {
            if (files.length > 1) {
                // This is a split file, try to merge it
                const mergedData = files.map(f => f.data).join("");
                const mergedFile = { name: fileName, data: mergedData };
                const testMergedList = [...mergedFiles, mergedFile];
                const testMergedJson = JSON.stringify(testMergedList);

                if (testMergedJson.length <= 1000000) {
                    // Merged file can fit in main field
                    mergedFiles.push(mergedFile);
                } else {
                    // Keep the split files
                    remainingFiles.push(...files);
                }
            } else {
                // Single file, add to appropriate list
                const testList = [...mergedFiles, files[0]];
                const testJson = JSON.stringify(testList);

                if (testJson.length <= 1000000) {
                    mergedFiles.push(files[0]);
                } else {
                    remainingFiles.push(files[0]);
                }
            }
        });

        // If all files can fit in main field, return them there
        if (remainingFiles.length === 0) {
            return {
                uploadedFile: JSON.stringify(mergedFiles),
                uploadedFileOverflow: undefined
            };
        }

        // Otherwise, distribute between main and overflow fields
        const mainFieldFiles: { name: string; data: string }[] = [];
        const overflowFiles: { name: string; data: string }[] = [];
        let currentLength = 0;

        // First add merged files to main field
        mergedFiles.forEach(file => {
            const fileJson = JSON.stringify(file);
            const testLength = currentLength + fileJson.length + 1;
            
            if (testLength <= 1000000) {
                mainFieldFiles.push(file);
                currentLength = testLength;
            } else {
                overflowFiles.push(file);
            }
        });

        // Then add remaining files
        remainingFiles.forEach(file => {
            const fileJson = JSON.stringify(file);
            const testLength = currentLength + fileJson.length + 1;
            
            if (testLength <= 1000000) {
                mainFieldFiles.push(file);
                currentLength = testLength;
            } else {
                overflowFiles.push(file);
            }
        });

        return {
            uploadedFile: JSON.stringify(mainFieldFiles),
            uploadedFileOverflow: JSON.stringify(overflowFiles)
        };
    }

    public destroy(): void {
        // Cleanup if necessary
    }
}
