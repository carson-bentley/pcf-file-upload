import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class AttachmentControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private fileInput: HTMLInputElement;
    private previewContainer: HTMLDivElement;
    private notifyOutputChanged: () => void;
    private fileBase64List: { name: string; data: string }[] = []; // List of uploaded files (name + base64)

    private readonly maxFileSizeMB: number = 5; // Maximum file size in MB
    private readonly acceptedFileTypes: string[] = ["image/", "application/pdf", "text/"]; // Allowed file types

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;

        // Check if there is any data in the uploadedFile field
        const uploadedFileData = context.parameters.uploadedFile.raw;
        if (uploadedFileData) {
            try {
                const parsedData = JSON.parse(uploadedFileData);
                this.fileBase64List = parsedData; // Populate the fileBase64List with the parsed data
            } catch (error) {
                console.error("Error parsing uploaded file data:", error);
            }
        }

        // Initialize the UI
        this.createUI();
        this.displayPreviews(); // Display the previews after initializing the UI
    }

    private createUI(): void {
        // File input (hidden, allows multiple files)
        this.fileInput = document.createElement("input");
        this.fileInput.type = "file";
        this.fileInput.accept = this.acceptedFileTypes.join(","); // Allowed file types
        this.fileInput.style.display = "none";
        this.fileInput.multiple = true; // Allow multiple files
        this.fileInput.onchange = this.handleFileUpload.bind(this);

        // Upload icon
        const uploadIcon = this.createUploadIcon();
        uploadIcon.onclick = () => this.fileInput.click();

        // Drag-and-drop area
        const dragDropArea = this.createDragDropArea();

        // Preview container
        this.previewContainer = document.createElement("div");
        this.previewContainer.style.border = "1px solid #ddd";
        this.previewContainer.style.padding = "10px";
        this.previewContainer.style.marginTop = "10px";
        this.previewContainer.textContent = "No files uploaded.";

        // Append elements to the container
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
            this.fileBase64List.push({ name: file.name, data: fileData });
            this.notifyOutputChanged();
            this.displayPreviews();
        };
        reader.readAsDataURL(file); // Convert file to Base64
    }

    private validateFile(file: File): boolean {
        // Check file size
        const maxSizeBytes = this.maxFileSizeMB * 1024 * 1024;
        if (file.size > maxSizeBytes) {
            alert(`File "${file.name}" exceeds the ${this.maxFileSizeMB} MB size limit.`);
            return false;
        }

        // Check file type
        if (!this.acceptedFileTypes.some((type) => file.type.startsWith(type))) {
            alert(`Unsupported file type for "${file.name}". Please upload a valid file.`);
            return false;
        }

        return true;
    }

    private createFullScreenPreview(): HTMLDivElement {
        // Create container at document level instead of component level
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
        closeButton.innerHTML = `
            <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                <line x1="18" y1="6" x2="6" y2="18"></line>
                <line x1="6" y1="6" x2="18" y2="18"></line>
            </svg>
        `;
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

        // Add click handler to close when clicking outside content
        fullScreenContainer.addEventListener("click", (event) => {
            if (event.target === fullScreenContainer) {
                fullScreenContainer.style.display = "none";
            }
        });

        // Add escape key handler
        const escapeHandler = (event: KeyboardEvent) => {
            if (event.key === "Escape") {
                fullScreenContainer.style.display = "none";
            }
        };
        document.addEventListener("keydown", escapeHandler);

        // Append to document body instead of component container
        document.body.appendChild(fullScreenContainer);

        return fullScreenContainer;
    }

    private displayPreviews(): void {
        this.previewContainer.innerHTML = ""; // Clear previous previews

        if (this.fileBase64List.length === 0) {
            this.previewContainer.textContent = "No files uploaded.";
            return;
        }

        // Create fullScreenContainer at document level
        const fullScreenContainer = this.createFullScreenPreview();

        this.fileBase64List.forEach((file, index) => {
            const fileContainer = document.createElement("div");
            fileContainer.style.marginBottom = "10px";
            fileContainer.style.position = "relative"; // Position relative for absolute positioning of the button

            const fileLabel = document.createElement("p");
            fileLabel.textContent = `File ${index + 1}: ${file.name}`;

            const removeButton = document.createElement("button");
            removeButton.innerHTML = `
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M3 6h18"></path>
                    <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"></path>
                    <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"></path>
                </svg>
            `;
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
            removeButton.onclick = () => this.removeFile(index);

            const fullScreenButton = document.createElement("button");
            fullScreenButton.innerHTML = `
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3"></path>
                </svg>
            `;
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
                    previewContent.innerHTML = ""; // Clear previous content
                    if (file.data.startsWith("data:image/")) {
                        const img = document.createElement("img");
                        img.src = file.data;
                        img.style.maxWidth = "100%";
                        previewContent.appendChild(img);
                    } else if (file.data.startsWith("data:application/pdf")) {
                        const iframe = document.createElement("iframe");
                        iframe.src = file.data;
                        iframe.style.width = "100%";
                        iframe.style.height = "100%";
                        previewContent.appendChild(iframe);
                    } else if (file.data.startsWith("data:text/")) {
                        const textArea = document.createElement("textarea");
                        textArea.value = atob(file.data.split(",")[1]); // Decode Base64
                        textArea.readOnly = true;
                        textArea.style.width = "100%";
                        textArea.style.height = "100%";
                        previewContent.appendChild(textArea);
                    }
                    fullScreenContainer.style.display = "block";
                }
            };

            fileContainer.appendChild(fileLabel);
            fileContainer.appendChild(removeButton);
            fileContainer.appendChild(fullScreenButton);

            if (file.data.startsWith("data:image/")) {
                const img = document.createElement("img");
                img.src = file.data;
                img.style.maxWidth = "100%";
                fileContainer.appendChild(img);
            } else if (file.data.startsWith("data:application/pdf")) {
                const iframe = document.createElement("iframe");
                iframe.src = file.data;
                iframe.style.width = "100%";
                iframe.style.height = "200px";
                fileContainer.appendChild(iframe);
            } else if (file.data.startsWith("data:text/")) {
                const textArea = document.createElement("textarea");
                textArea.value = atob(file.data.split(",")[1]); // Decode Base64
                textArea.readOnly = true;
                textArea.style.width = "100%";
                textArea.style.height = "100px";
                fileContainer.appendChild(textArea);
            }

            this.previewContainer.appendChild(fileContainer);
        });
    }

    private removeFile(index: number): void {
        this.fileBase64List.splice(index, 1); // Remove the file from the list
        this.notifyOutputChanged();
        this.displayPreviews();
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Update the view when bound data changes
    }

    public getOutputs(): IOutputs {
        const fileData = this.fileBase64List.length > 0 ? this.fileBase64List[0].data : null;
        return { uploadedFile: JSON.stringify(this.fileBase64List) }; // Return files as a JSON string
        
    
    }

    public destroy(): void {
        // Cleanup if necessary
    }
}
