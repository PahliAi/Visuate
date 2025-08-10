/**
 * Translation Manager for Equate Portfolio Analysis
 * Handles multi-language support for 13 languages from EquatePlus
 */

class TranslationManager {
    constructor() {
        this.currentLanguage = 'english';
        this.detectedLanguage = null;
        this.translations = {};
        this.languageMap = {
            'english': 'english',
            'german': 'german', 
            'dutch': 'dutch',
            'french': 'french',
            'spanish': 'spanish',
            'italian': 'italian',
            'polish': 'polish',
            'turkish': 'turkish',
            'portuguese': 'portuguese',
            'czech': 'czech',
            'romanian': 'romanian',
            'croatian': 'croatian',
            'indonesian': 'indonesian',
            'chinese': 'chinese'
        };
        
        // Language detection patterns - key terms that uniquely identify each language
        this.languagePatterns = {
            german: {
                required: ['Benutzerkennung', 'Allokationsdatum', 'Portfolioangaben'],
                optional: ['Ausführungspreis', 'Marktpreis', 'Verfügbar ab']
            },
            dutch: {
                required: ['Gebruikers-ID', 'Allocatiedatum', 'Portefeuillegegevens'],
                optional: ['Uitvoeringsprijs', 'Marktprijs', 'Beschikbaar vanaf']
            },
            french: {
                required: ['Identifiant d\'utilisateur', 'Date de l\'allocation', 'Détails du portefeuille'],
                optional: ['Prix d\'exécution', 'Prix de marché', 'Disponible à partir de']
            },
            spanish: {
                required: ['Identificador de usuario', 'Fecha de asignación', 'Detalles de la cartera'],
                optional: ['Precio de ejecución', 'Precio de mercado', 'Disponible desde']
            },
            italian: {
                required: ['ID utente', 'Data di allocazione', 'Dettagli del portafoglio'],
                optional: ['Prezzo di esecuzione', 'Prezzo di mercato', 'Disponibile da']
            },
            polish: {
                required: ['Numer ID użytkownika', 'Data alokacji', 'Szczegóły dotyczące portfolio'],
                optional: ['Cena wykonania', 'Cena rynkowa', 'Dostępne od']
            },
            turkish: {
                required: ['Kullanıcı kimliği', 'Tahsis tarihi', 'Portföy bilgileri'],
                optional: ['İşlem fiyatı', 'Piyasa fiyatı', 'Kullanıma sunulacağı tarih']
            },
            portuguese: {
                required: ['Nome de usuário', 'Data de alocação', 'Detalhes do portfólio'],
                optional: ['Preço de execução', 'Preço de mercado', 'Disponível a partir de']
            },
            czech: {
                required: ['Uživatelské ID', 'Datum alokace', 'Jméno účastníka'], // Actual Czech terms from files
                optional: ['Realizační cena', 'Tržní cena', 'Reference pokynu', 'Typ pokynu']
            },
            romanian: {
                required: ['ID utilizator', 'Dată de alocare', 'Detalii portofoliu'],
                optional: ['Preț de executare', 'Preț de piață', 'Disponibil de la']
            },
            croatian: {
                required: ['Korisnički ID', 'Datum alociranja', 'Pojedinosti portfelja'],
                optional: ['Cijena izvršavanja', 'Tržišna cijena', 'Dostupno od']
            },
            indonesian: {
                required: ['ID Pengguna', 'Tanggal alokasi', 'Detail portofolio'],
                optional: ['Harga pelaksanaan', 'Harga pasar', 'Tersedia dari']
            },
            chinese: {
                required: ['用户ID', '截至日期', '分配日期'],
                optional: ['计划', '供款类型', '订单参考', '订单类型', '执行价格', '数量', '状态']
            },
            english: {
                required: ['User ID', 'Allocation date', 'Portfolio Details'], // Default fallback
                optional: ['Execution price', 'Market price', 'Available from']
            }
        };
    }

    /**
     * Initialize translations from embedded data
     */
    async init() {
        try {
            // Load translations from embedded TRANSLATION_DATA
            this.loadTranslations();
            config.debug('✅ Translation Manager initialized with', Object.keys(TRANSLATION_DATA).length, 'translation keys');
        } catch (error) {
            config.error('❌ Failed to initialize Translation Manager:', error);
            // Fallback to English-only mode
            this.translations = { english: {} };
        }
    }

    /**
     * Load translations from embedded TRANSLATION_DATA
     */
    loadTranslations() {
        // Convert TRANSLATION_DATA format to internal format
        this.translations = {};
        
        // Initialize language objects
        const languages = ['english', 'german', 'dutch', 'french', 'spanish', 'italian', 
                          'polish', 'turkish', 'portuguese', 'czech', 'romanian', 'croatian', 'indonesian', 'chinese'];
        
        languages.forEach(lang => {
            this.translations[lang] = {};
        });
        
        // Copy translations from TRANSLATION_DATA
        Object.keys(TRANSLATION_DATA).forEach(key => {
            const translations = TRANSLATION_DATA[key];
            languages.forEach(lang => {
                if (translations[lang]) {
                    this.translations[lang][key] = translations[lang];
                }
            });
        });
        
        config.debug('🌐 Loaded translations for languages:', Object.keys(this.translations));
        config.debug('🔑 Sample translations loaded:', {
            your_investment_title: this.translations.english.your_investment_title,
            company_match_label: this.translations.english.company_match_label
        });
    }

    /**
     * Detect language from Excel file content
     */
    async detectLanguageFromFile(file) {
        try {
            // Read Excel file
            const data = await this.readExcelFile(file);
            const worksheet = data.Sheets[data.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                raw: true
            });

            // Extract all text content from first 20 rows
            const textContent = [];
            for (let i = 0; i < Math.min(20, rawData.length); i++) {
                const row = rawData[i];
                if (row) {
                    row.forEach(cell => {
                        if (typeof cell === 'string' && cell.trim()) {
                            textContent.push(cell.trim());
                        }
                    });
                }
            }

            // Score each language based on pattern matches
            const languageScores = {};
            
            Object.keys(this.languagePatterns).forEach(lang => {
                languageScores[lang] = 0;
                const patterns = this.languagePatterns[lang];
                
                // Check required patterns (high weight)
                patterns.required.forEach(pattern => {
                    if (textContent.some(text => text.includes(pattern))) {
                        languageScores[lang] += 10;
                    }
                });
                
                // Check optional patterns (low weight)
                patterns.optional.forEach(pattern => {
                    if (textContent.some(text => text.includes(pattern))) {
                        languageScores[lang] += 2;
                    }
                });
            });

            // Find language with highest score
            let detectedLang = 'english'; // Default
            let maxScore = 0;
            
            Object.keys(languageScores).forEach(lang => {
                if (languageScores[lang] > maxScore) {
                    maxScore = languageScores[lang];
                    detectedLang = lang;
                }
            });

            // Require minimum score to avoid false positives
            if (maxScore < 10) {
                config.warn('⚠️ No strong language detection match, defaulting to English');
                detectedLang = 'english';
            }

            this.detectedLanguage = detectedLang;
            config.debug('🌐 Language detected:', detectedLang, 'with score:', maxScore);
            config.debug('🔍 Language scores:', languageScores);
            config.debug('📝 Sample text content:', textContent.slice(0, 10));

            return detectedLang;
        } catch (error) {
            config.error('❌ Language detection failed:', error);
            return 'english'; // Safe fallback
        }
    }

    /**
     * Read Excel file using SheetJS (reused from FileParser)
     */
    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { 
                        type: 'array',
                        cellDates: true,
                        raw: false
                    });
                    resolve(workbook);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = function() {
                reject(new Error('Failed to read file'));
            };
            
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Get translation for a key in current language
     */
    t(key, fallback = null) {
        const lang = this.currentLanguage;
        
        if (this.translations[lang] && this.translations[lang][key]) {
            return this.translations[lang][key];
        }
        
        // Fallback to English
        if (lang !== 'english' && this.translations.english && this.translations.english[key]) {
            return this.translations.english[key];
        }
        
        // Return fallback or key as last resort
        const result = fallback || key;
        
        // Debug logging for missing translations
        if (result === key) {
            config.warn('⚠️ Missing translation for key:', key, 'in language:', lang);
        }
        
        return result;
    }

    /**
     * Set current language and update UI
     */
    setLanguage(language) {
        if (this.languageMap[language]) {
            this.currentLanguage = language;
            config.debug('🌐 Language set to:', language);
            
            // Update UI elements
            this.updateUILanguage();
        } else {
            config.warn('⚠️ Unsupported language:', language);
        }
    }

    /**
     * Update UI elements with current language
     */
    updateUILanguage() {
        // Update page title
        document.title = this.t('page_title', 'Equate - Portfolio Analysis');
        
        // Update main navigation tabs
        const howtoTab = document.getElementById('howtoTab');
        if (howtoTab) howtoTab.innerHTML = '📋 ' + this.t('howto_tab', 'How to');
        
        const uploadTab = document.getElementById('uploadTab');
        if (uploadTab) uploadTab.innerHTML = '📤 ' + this.t('upload_tab', 'Upload');
        
        const resultsTab = document.getElementById('resultsTab');
        if (resultsTab) resultsTab.innerHTML = '📊 ' + this.t('results_tab', 'Results');

        // Update file upload labels
        const portfolioZone = document.querySelector('#portfolioZone h3');
        if (portfolioZone) portfolioZone.textContent = this.t('portfolio_details_title');
        
        const portfolioDesc = document.querySelector('#portfolioZone p');
        if (portfolioDesc) portfolioDesc.innerHTML = this.t('drop_portfolio_file');
        
        const transactionZone = document.querySelector('#transactionsZone h3');
        if (transactionZone) transactionZone.textContent = this.t('transaction_history_title');
        
        const transactionDesc = document.querySelector('#transactionsZone p');
        if (transactionDesc) transactionDesc.innerHTML = this.t('drop_transaction_file');

        // Update buttons
        const analyzeBtn = document.getElementById('analyzeBtn');
        if (analyzeBtn && !analyzeBtn.disabled) {
            analyzeBtn.textContent = this.t('analyze_button');
        }
        
        const clearBtn = document.querySelector('.clear-data-btn');
        if (clearBtn) clearBtn.innerHTML = '🗑️ ' + this.t('clear_data_button');
        
        const exportBtn = document.getElementById('resultsExportBtn');
        if (exportBtn) exportBtn.innerHTML = '📋 ' + this.t('export_pdf_button');
        
        const updatePriceBtn = document.querySelector('.quick-btn');
        if (updatePriceBtn) updatePriceBtn.textContent = this.t('update_price_button');

        // Update section order label
        const sectionOrderLabel = document.getElementById('sectionOrderLabel');
        if (sectionOrderLabel) sectionOrderLabel.textContent = this.t('section_order_label') + ':';

        // Update metric labels
        this.updateMetricLabels();
        
        // Update chart titles
        this.updateChartTitles();
        
        // Update price source indicator if it exists
        this.updatePriceSourceIndicator();
        
        // Update privacy footer
        this.updatePrivacyFooter();
        
        // Update timeline disclaimer if it's currently displayed  
        this.refreshTimelineDisclaimerLanguage();
    }

    /**
     * Update metric card labels
     */
    updateMetricLabels() {
        const metricUpdates = [
            // Row 1: Your Investment, Company Match, Free Shares, Dividend Income, Total Investment
            { selector: '.metric-card h3', key: 'your_investment_title', index: 0 },
            { selector: '.metric-card h3', key: 'company_match_label', index: 1 },
            { selector: '.metric-card h3', key: 'free_shares_label', index: 2 },
            { selector: '.metric-card h3', key: 'dividend_income', index: 3 },
            { selector: '.metric-card h3', key: 'total_investment', index: 4 },
            // Row 2: Return on Your Investment, Current Portfolio, Total Sold, Total Value, Return on Total Investment
            { selector: '.metric-card h3', key: 'return_on_investment', index: 5 },
            { selector: '.metric-card h3', key: 'current_portfolio', index: 6 },
            { selector: '.metric-card h3', key: 'total_sold', index: 7 },
            { selector: '.metric-card h3', key: 'total_value', index: 8 },
            { selector: '.metric-card h3', key: 'return_on_total_investment', index: 9 },
            // Row 3: Return % on Your Investment, Annual Return XIRR on Your Investment, Available Shares, Annual Return XIRR on Total Investment, Return % on Total Investment
            { selector: '.metric-card h3', key: 'return_percentage_on_investment', index: 10 },
            { selector: '.metric-card h3', key: 'xirr_user_investment', index: 11 },
            { selector: '.metric-card h3', key: 'available_shares_label', index: 12 },
            { selector: '.metric-card h3', key: 'xirr_total_investment', index: 13 },
            { selector: '.metric-card h3', key: 'return_percentage_on_total_investment', index: 14 }
        ];

        metricUpdates.forEach(update => {
            const elements = document.querySelectorAll(update.selector);
            if (elements[update.index]) {
                const translation = this.t(update.key);
                // Debug logging
                // config.debug('🌐 Translating metric', update.index, 'from key', update.key, 'to:', translation);
                elements[update.index].textContent = translation;
            }
        });
    }

    /**
     * Update chart titles
     */
    updateChartTitles() {
        const chartTitles = document.querySelectorAll('.chart-container h3');
        if (chartTitles[0]) chartTitles[0].textContent = this.t('portfolio_timeline_title');
        if (chartTitles[1]) chartTitles[1].textContent = this.t('performance_overview_title');
        if (chartTitles[2]) chartTitles[2].textContent = this.t('investment_sources_title');
    }

    /**
     * Update price source indicator
     */
    updatePriceSourceIndicator() {
        // Trigger price source update if the app instance exists
        if (typeof app !== 'undefined' && app.currentCalculations) {
            app.updatePriceSource();
        }
    }

    /**
     * Get column header translation for file parsing
     */
    getColumnHeader(key) {
        return this.t(key);
    }

    /**
     * Get supported languages list
     */
    getSupportedLanguages() {
        return Object.keys(this.languageMap);
    }

    /**
     * Get current language
     */
    getCurrentLanguage() {
        return this.currentLanguage;
    }

    /**
     * Get detected language from file analysis
     */
    getDetectedLanguage() {
        return this.detectedLanguage;
    }

    /**
     * Create language selector UI element
     */
    createLanguageSelector() {
        const selector = document.createElement('div');
        selector.className = 'language-selector';
        selector.innerHTML = `
            <label for="languageSelect">${this.t('language_label', 'Language')}:</label>
            <select id="languageSelect" onchange="translationManager.setLanguage(this.value)">
                ${this.getSupportedLanguages().map(lang => 
                    `<option value="${lang}" ${lang === this.currentLanguage ? 'selected' : ''}>
                        ${this.t('lang_' + lang, lang.charAt(0).toUpperCase() + lang.slice(1))}
                    </option>`
                ).join('')}
            </select>
        `;
        return selector;
    }
    
    /**
     * Update privacy footer text
     */
    updatePrivacyFooter() {
        const privacyFooter = document.querySelector('.privacy-footer');
        if (privacyFooter) {
            privacyFooter.innerHTML = `
                <strong>${this.t('privacy_statement')}</strong> ${this.t('privacy_description')}
                <div style="margin-top: 10px; font-size: 12px; opacity: 0.7;">
                    ${this.t('app_version_info')}
                </div>
            `;
        }
    }
    
    /**
     * Update timeline disclaimer and total sold card disclaimer language when language changes
     */
    refreshTimelineDisclaimerLanguage() {
        config.debug('🌐 refreshTimelineDisclaimerLanguage called');
        
        // Check if the app instance exists and has a method to refresh the disclaimers
        if (typeof app !== 'undefined' && typeof app.refreshTimelineDisclaimer === 'function') {
            // Call the app's method to refresh both disclaimers with current language
            app.refreshTimelineDisclaimer();
            config.debug('🌐 Called app.refreshTimelineDisclaimer() for language change (both timeline and card)');
        } else {
            config.debug('⚠️ app.refreshTimelineDisclaimer not available for language update', {
                appExists: typeof app !== 'undefined',
                methodExists: typeof app !== 'undefined' ? typeof app.refreshTimelineDisclaimer : 'N/A'
            });
        }
    }
}

// Global translation manager instance
const translationManager = new TranslationManager();