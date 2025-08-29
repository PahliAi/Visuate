/**
 * Export Customization Module for Equate Portfolio Analysis
 * Handles PDF export section selection and ordering
 */

class ExportCustomizer {
    constructor() {
        this.modal = null;
        this.sectionList = null;
        this.selectedIndex = -1;
        
        // Default sections configuration with 13 total sections
        this.defaultSections = [
            // Main Report Sections (4 items)
            { id: 'portfolioCards', name: 'Portfolio Cards', category: 'main', enabled: true },
            { id: 'timelineChart', name: 'Timeline Chart', category: 'main', enabled: true },
            { id: 'performanceBarChart', name: 'Performance Bar Chart', category: 'main', enabled: true },
            { id: 'investmentPieChart', name: 'Investment Pie Chart', category: 'main', enabled: true },
            
            // Calculation Breakdown Sections (9 items)
            { id: 'yourInvestment', name: 'Your Investment', category: 'calculation', enabled: true },
            { id: 'companyMatch', name: 'Company Match', category: 'calculation', enabled: true },
            { id: 'freeShares', name: 'Free Shares', category: 'calculation', enabled: true },
            { id: 'dividendIncome', name: 'Dividend Income', category: 'calculation', enabled: true },
            { id: 'totalInvestment', name: 'Total Investment', category: 'calculation', enabled: true },
            { id: 'currentPortfolio', name: 'Current Portfolio', category: 'calculation', enabled: true },
            { id: 'totalSold', name: 'Total Sold', category: 'calculation', enabled: true },
            { id: 'xirrYourInvestment', name: 'XIRR - Your Investment', category: 'calculation', enabled: true },
            { id: 'xirrTotalInvestment', name: 'XIRR - Total Investment', category: 'calculation', enabled: true }
        ];

        // Initialize from localStorage or defaults
        this.sections = this.loadSectionPreferences();
    }

    /**
     * Initialize the export customizer
     */
    init() {
        // Don't create modal on init - create it when first shown
        // this.createModal();
        // this.bindEvents();
    }

    /**
     * Show the customization modal
     */
    show() {
        if (!this.modal) {
            this.createModal();
            this.bindEvents(); // Bind events after modal is created
            this.applyTranslations(); // Apply translations after modal is created
        }
        
        this.refreshSectionList();
        this.modal.style.display = 'block';
        
        // Focus on modal for accessibility
        setTimeout(() => {
            const firstButton = this.modal.querySelector('.modal-controls button');
            if (firstButton) {
                firstButton.focus();
            }
        }, 100);
    }

    /**
     * Hide the customization modal
     */
    hide() {
        if (this.modal) {
            this.modal.style.display = 'none';
        }
        this.selectedIndex = -1;
    }

    /**
     * Create the modal HTML structure
     */
    createModal() {
        const modalHTML = `
            <div id="exportCustomizerModal" class="export-modal">
                <div class="modal-content">
                    <div class="modal-header">
                        <h2 data-translate="export.customize.title">Customize PDF Export</h2>
                        <button class="modal-close" aria-label="Close">&times;</button>
                    </div>
                    
                    <div class="modal-body">
                        <div class="section-categories-sidebyside">
                            <div class="category-column">
                                <h3 data-translate="export.customize.mainSections">Main Report Sections (4 items)</h3>
                                <div id="mainSectionsList" class="sections-list-column"></div>
                            </div>
                            
                            <div class="category-column">
                                <h3 data-translate="export.customize.calculationSections">Calculation Breakdown Sections (9 items)</h3>
                                <div id="calculationSectionsList" class="sections-list-column"></div>
                            </div>
                        </div>

                        <div class="modal-controls">
                            <div class="move-controls-center">
                                <div class="move-buttons-fixed">
                                    <button id="moveUpBtn" class="move-control-btn" disabled>
                                        <span aria-hidden="true">↑</span>
                                        <span data-translate="export.customize.moveUp">Move Up</span>
                                    </button>
                                    <button id="moveDownBtn" class="move-control-btn" disabled>
                                        <span aria-hidden="true">↓</span>
                                        <span data-translate="export.customize.moveDown">Move Down</span>
                                    </button>
                                </div>
                            </div>
                            
                            <div class="control-group">
                                <button id="enableAllBtn" class="control-btn">
                                    <span data-translate="export.customize.enableAll">Enable All</span>
                                </button>
                                <button id="disableAllBtn" class="control-btn">
                                    <span data-translate="export.customize.disableAll">Disable All</span>
                                </button>
                                <button id="resetOrderBtn" class="control-btn">
                                    <span data-translate="export.customize.resetOrder">Reset Order</span>
                                </button>
                            </div>
                            
                        </div>
                    </div>
                    
                    <div class="modal-footer">
                        <button id="cancelExportBtn" class="btn-secondary">
                            <span data-translate="export.customize.cancel">Cancel</span>
                        </button>
                        <button id="generateCustomPdfBtn" class="btn-primary">
                            <span data-translate="export.customize.generatePdf">Generate Custom PDF</span>
                        </button>
                    </div>
                </div>
            </div>
        `;

        // Add modal to DOM
        document.body.insertAdjacentHTML('beforeend', modalHTML);
        this.modal = document.getElementById('exportCustomizerModal');
        this.sectionList = this.modal.querySelector('.sections-list');
    }

    /**
     * Bind event listeners
     */
    bindEvents() {
        if (!this.modal) return;

        // Close modal events
        this.modal.querySelector('.modal-close').addEventListener('click', () => this.hide());
        this.modal.querySelector('#cancelExportBtn').addEventListener('click', () => this.hide());
        
        // Move buttons
        this.modal.querySelector('#moveUpBtn').addEventListener('click', () => this.moveSelectedSection(-1));
        this.modal.querySelector('#moveDownBtn').addEventListener('click', () => this.moveSelectedSection(1));
        
        // Control buttons
        this.modal.querySelector('#enableAllBtn').addEventListener('click', () => this.enableAllSections());
        this.modal.querySelector('#disableAllBtn').addEventListener('click', () => this.disableAllSections());
        this.modal.querySelector('#resetOrderBtn').addEventListener('click', () => this.resetOrder());
        
        // Generate PDF button
        this.modal.querySelector('#generateCustomPdfBtn').addEventListener('click', () => this.generateCustomPDF());
        
        // Close modal when clicking outside
        this.modal.addEventListener('click', (e) => {
            if (e.target === this.modal) {
                this.hide();
            }
        });

        // No scroll handling needed - modal fits all content

        // Keyboard navigation
        this.modal.addEventListener('keydown', (e) => this.handleKeydown(e));
    }

    /**
     * Handle keyboard navigation
     */
    handleKeydown(e) {
        if (e.key === 'Escape') {
            this.hide();
        } else if (e.key === 'ArrowUp' && this.selectedIndex >= 0) {
            e.preventDefault();
            this.moveSelectedSection(-1);
        } else if (e.key === 'ArrowDown' && this.selectedIndex >= 0) {
            e.preventDefault();
            this.moveSelectedSection(1);
        }
    }

    /**
     * Refresh the sections list display
     */
    refreshSectionList() {
        const mainList = this.modal.querySelector('#mainSectionsList');
        const calcList = this.modal.querySelector('#calculationSectionsList');
        
        mainList.innerHTML = '';
        calcList.innerHTML = '';

        this.sections.forEach((section, index) => {
            const sectionElement = this.createSectionElement(section, index);
            
            if (section.category === 'main') {
                mainList.appendChild(sectionElement);
            } else {
                calcList.appendChild(sectionElement);
            }
        });

        // No floating buttons to hide - using fixed buttons
    }

    /**
     * Create a section element
     */
    createSectionElement(section, index) {
        const element = document.createElement('div');
        element.className = `section-item ${section.enabled ? 'enabled' : 'disabled'}`;
        element.dataset.index = index;
        
        element.innerHTML = `
            <div class="section-order">${index + 1}</div>
            <div class="section-checkbox">
                <span class="checkbox-icon">${section.enabled ? '✓' : '✗'}</span>
            </div>
            <div class="section-content">
                <div class="section-name ${!section.enabled ? 'crossed-out' : ''}">${section.name}</div>
            </div>
        `;

        // Click events
        element.addEventListener('click', (e) => {
            if (!e.target.closest('.section-checkbox')) {
                this.selectSection(index);
            }
        });
        
        element.querySelector('.section-checkbox').addEventListener('click', (e) => {
            e.stopPropagation();
            this.toggleSection(index);
        });

        return element;
    }

    /**
     * Get description for a section
     */
    getSectionDescription(sectionId) {
        const descriptions = {
            'portfolioCards': 'Key metrics summary with investment breakdown and returns',
            'timelineChart': 'Portfolio performance over time with transaction markers',
            'performanceBarChart': 'Stacked bar chart comparing investment vs returns',
            'investmentPieChart': 'Distribution breakdown of investment sources',
            'yourInvestment': 'Breakdown of your personal contributions and current value',
            'companyMatch': 'Employer contribution breakdown and current value',
            'freeShares': 'Free share awards and their current market value',
            'dividendIncome': 'Dividend reinvestments and accumulated value',
            'totalInvestment': 'Summary table of all investment sources and percentages',
            'currentPortfolio': 'Current holdings with profit/loss analysis',
            'totalSold': 'Historical sales with prices and proceeds',
            'xirrYourInvestment': 'Annual return calculation with cash flow timeline',
            'xirrTotalInvestment': 'Annual return for total investment with cash flow analysis'
        };
        return descriptions[sectionId] || '';
    }

    /**
     * Select a section for moving
     */
    selectSection(index) {
        // Remove previous selection
        this.modal.querySelectorAll('.section-item').forEach(el => el.classList.remove('selected'));
        
        // Add selection to clicked item
        const selectedElement = this.modal.querySelector(`[data-index="${index}"]`);
        if (selectedElement) {
            selectedElement.classList.add('selected');
            this.selectedIndex = index;
            
        }
        
        this.updateMoveButtons();
    }


    /**
     * Toggle section enabled/disabled state
     */
    toggleSection(index) {
        this.sections[index].enabled = !this.sections[index].enabled;
        this.refreshSectionList();
        this.saveSectionPreferences();
    }

    /**
     * Move the currently selected section up or down
     */
    moveSelectedSection(direction) {
        if (this.selectedIndex < 0) return;
        
        const index = this.selectedIndex;
        const newIndex = index + direction;
        
        // Check bounds within the same category
        const currentSection = this.sections[index];
        const categorySections = this.sections.filter(s => s.category === currentSection.category);
        const categoryStartIndex = this.sections.findIndex(s => s.category === currentSection.category);
        const categoryEndIndex = categoryStartIndex + categorySections.length - 1;
        
        // Check if move is within category bounds
        if (newIndex < categoryStartIndex || newIndex > categoryEndIndex) {
            return; // Can't move outside category bounds
        }
        
        // Swap sections
        [this.sections[index], this.sections[newIndex]] = 
        [this.sections[newIndex], this.sections[index]];
        
        // Update selectedIndex to follow the moved item
        this.selectedIndex = newIndex;
        
        this.refreshSectionList();
        this.saveSectionPreferences();
        
        // Re-select the moved item at its new position
        setTimeout(() => this.selectSection(newIndex), 10);
    }

    /**
     * Update move button states based on current selection
     */
    updateMoveButtons() {
        const moveUpBtn = this.modal.querySelector('#moveUpBtn');
        const moveDownBtn = this.modal.querySelector('#moveDownBtn');
        
        if (this.selectedIndex < 0) {
            moveUpBtn.disabled = true;
            moveDownBtn.disabled = true;
            return;
        }
        
        const currentSection = this.sections[this.selectedIndex];
        const categorySections = this.sections.filter(s => s.category === currentSection.category);
        const categoryStartIndex = this.sections.findIndex(s => s.category === currentSection.category);
        const localIndex = this.selectedIndex - categoryStartIndex;
        
        moveUpBtn.disabled = localIndex === 0;
        moveDownBtn.disabled = localIndex === categorySections.length - 1;
    }

    /**
     * Enable all sections
     */
    enableAllSections() {
        this.sections.forEach(section => section.enabled = true);
        this.refreshSectionList();
        this.saveSectionPreferences();
    }

    /**
     * Disable all sections
     */
    disableAllSections() {
        this.sections.forEach(section => section.enabled = false);
        this.refreshSectionList();
        this.saveSectionPreferences();
    }

    /**
     * Reset to default order
     */
    resetOrder() {
        this.sections = JSON.parse(JSON.stringify(this.defaultSections));
        this.selectedIndex = -1;
        this.refreshSectionList();
        this.saveSectionPreferences();
    }


    /**
     * Generate custom PDF with selected sections
     */
    async generateCustomPDF() {
        const enabledSections = this.sections.filter(section => section.enabled);
        
        if (enabledSections.length === 0) {
            alert('Please select at least one section to include in the PDF.');
            return;
        }
        
        // Save preferences
        this.saveSectionPreferences();
        
        // Close modal
        this.hide();
        
        // Generate PDF with custom configuration
        const exportConfig = {
            sections: enabledSections,
            sectionOrder: this.sections.map(s => s.id)
        };
        
        try {
            // Call the main export function with our configuration
            if (typeof portfolioExporter !== 'undefined') {
                await portfolioExporter.exportCustomPDF(exportConfig);
            } else {
                console.error('Portfolio exporter not available');
                alert('Portfolio exporter is not available. Please make sure the page is fully loaded.');
            }
        } catch (error) {
            console.error('Error generating custom PDF:', error);
            alert('Error generating PDF. Please try again.');
        }
    }

    /**
     * Get current export configuration
     */
    getExportConfig() {
        return {
            sections: this.sections.filter(s => s.enabled),
            sectionOrder: this.sections.map(s => s.id),
            allSections: this.sections
        };
    }

    /**
     * Load section preferences from localStorage
     */
    loadSectionPreferences() {
        try {
            const saved = localStorage.getItem('exportSectionPreferences');
            if (saved) {
                const parsed = JSON.parse(saved);
                // Preserve the saved order, but ensure we have all sections
                const result = [];
                
                // First, add all saved sections in their saved order
                parsed.forEach(savedSection => {
                    const defaultSection = this.defaultSections.find(d => d.id === savedSection.id);
                    if (defaultSection) {
                        result.push(savedSection);
                    }
                });
                
                // Then add any new sections that weren't in saved preferences
                this.defaultSections.forEach(defaultSection => {
                    if (!parsed.find(s => s.id === defaultSection.id)) {
                        result.push(defaultSection);
                    }
                });
                
                return result;
            }
        } catch (error) {
            console.warn('Error loading export preferences:', error);
        }
        
        return JSON.parse(JSON.stringify(this.defaultSections));
    }

    /**
     * Save section preferences to localStorage
     */
    saveSectionPreferences() {
        try {
            localStorage.setItem('exportSectionPreferences', JSON.stringify(this.sections));
        } catch (error) {
            console.warn('Error saving export preferences:', error);
        }
    }

    /**
     * Apply translations to modal (called after translation system loads)
     */
    applyTranslations() {
        if (!this.modal || typeof translationManager === 'undefined') return;
        
        const elementsToTranslate = this.modal.querySelectorAll('[data-translate]');
        elementsToTranslate.forEach(element => {
            const key = element.getAttribute('data-translate');
            const translation = translationManager.t(key);
            if (translation && translation !== key) {
                element.textContent = translation;
            }
        });
    }
}

// Global instance
window.exportCustomizer = new ExportCustomizer();