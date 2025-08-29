/**
 * INVESTMENT CATEGORIZATION - HYBRID RULES SYSTEM V3
 * Replaces hardcoded filter conditions with flexible regex patterns and company overrides
 */

const INVESTMENT_RULES = {
    // Base patterns applied to all companies (regex-based)
    BASE_PATTERNS: {
        userInvestment: {
            contributionType: /^Purchase$/i,
            plan: /Employee Share Purchase Plan$/i,  // Matches "IBM Employee Share Purchase Plan", "Allianz Employee Share Purchase Plan"
            instrument: null,         // Don't check - any instrument
            instrumentType: null,     // Don't check - any type
            orderType: null,          // Don't check - portfolio purchases don't need transaction fields
            status: null              // Don't check - portfolio purchases don't need transaction fields
        },
        companyMatch: {
            contributionType: /^Company match$/i,
            plan: /Employee Share Purchase Plan$/i,
            instrument: null,
            instrumentType: null,
            orderType: null,
            status: null
        },
        freeShares: {
            contributionType: /^Award$/i,
            plan: /^Free Share$/i,
            instrument: null,
            instrumentType: null,
            orderType: null,
            status: null
        },
        dividendIncome: {
            plan: /Dividend Reinvestment$/i,  // Matches "IBM Dividend Reinvestment", "Allianz Dividend Reinvestment"
            contributionType: null,
            instrument: null,
            instrumentType: null,
            orderType: null,
            status: null
        }
    },
    
    // Company-specific exact overrides (exact string matching)
    COMPANY_OVERRIDES: {
        'Nike': {  // Company-specific rules
            freeShares: {
                instrument: 'Nike Share',        // Must match Nike shares only
                plan: 'Greatness award',         // Exact match for Nike's special case
                contributionType: 'Award',       // Exact match
                instrumentType: null,            // Don't check type
                orderType: null,                 // Don't check transaction fields
                status: null                     // Don't check transaction fields
            }
        },
        // Future: 'IBM': { ... }, 'Allianz': { ... }
        // Test company for edge case validation
        'Test': {
            freeShares: {
                contributionType: 'Award',
                plan: 'Greatness award',
                instrument: null,
                instrumentType: null,
                orderType: null,
                status: null
            }
        }
    }
};

/**
 * Check if an entry matches a specific rule
 * @param {Object} entry - The data entry to check
 * @param {Object} rule - The rule to match against
 * @returns {boolean} - True if entry matches the rule
 */
function matchesRule(entry, rule) {
    // Helper function to check a field - if rule field is null, skip the check
    const checkField = (ruleValue, entryValue) => {
        if (ruleValue === null || ruleValue === undefined) {
            return true; // Skip check if rule doesn't specify this field
        }
        if (ruleValue instanceof RegExp) {
            return ruleValue.test(entryValue || '');
        }
        return entryValue === ruleValue;
    };

    return checkField(rule.contributionType, entry.contributionType) &&
           checkField(rule.plan, entry.plan) &&
           checkField(rule.instrument, entry.instrument) &&
           checkField(rule.instrumentType, entry.instrumentType) &&
           checkField(rule.orderType, entry.orderType) &&
           checkField(rule.status, entry.status);
}

/**
 * Categorize an entry using hybrid rule system
 * @param {Object} entry - The enhanced reference point entry
 * @param {string} company - Optional company name for overrides
 * @returns {string} - Category: 'userInvestment', 'companyMatch', 'freeShares', 'dividendIncome', or 'unknown'
 */
function categorizeEntry(entry, company = null) {
    // 1. Try company-specific overrides first (highest priority)
    if (company && INVESTMENT_RULES.COMPANY_OVERRIDES[company]) {
        const companyRules = INVESTMENT_RULES.COMPANY_OVERRIDES[company];
        for (const [category, rule] of Object.entries(companyRules)) {
            if (matchesRule(entry, rule)) {
                return category;
            }
        }
    }
    
    // 2. Fall back to base regex patterns
    const basePatterns = INVESTMENT_RULES.BASE_PATTERNS;
    for (const [category, rule] of Object.entries(basePatterns)) {
        if (matchesRule(entry, rule)) {
            return category;
        }
    }
    
    return 'unknown';
}

// Export for module systems
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { INVESTMENT_RULES, matchesRule, categorizeEntry };
} else {
    // For browser/ES6 modules
    window.INVESTMENT_RULES = INVESTMENT_RULES;
    window.matchesRule = matchesRule;
    window.categorizeEntry = categorizeEntry;
}