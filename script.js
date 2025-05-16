'use strict';

(function() {
    // --- Constants ---
    const COLUMN_EMAIL = 'Email';
    const COLUMN_USER_ID = 'User ID';
    const COLUMN_IS_ACTIVE = 'Is Active';
    const COLUMN_DATE = 'Date';
    const COLUMN_CHAT_SUGGESTED_LINES_ADDED = 'Chat Suggested Lines Added';
    const COLUMN_CHAT_SUGGESTED_LINES_DELETED = 'Chat Suggested Lines Deleted';
    const COLUMN_CHAT_ACCEPTED_LINES_ADDED = 'Chat Accepted Lines Added';
    const COLUMN_CHAT_ACCEPTED_LINES_DELETED = 'Chat Accepted Lines Deleted';
    const COLUMN_CHAT_TOTAL_APPLIES = 'Chat Total Applies';
    const COLUMN_CHAT_TOTAL_ACCEPTS = 'Chat Total Accepts';
    const COLUMN_CHAT_TOTAL_REJECTS = 'Chat Total Rejects';
    const COLUMN_CHAT_TABS_SHOWN = 'Chat Tabs Shown';
    const COLUMN_TABS_ACCEPTED = 'Tabs Accepted';
    const COLUMN_EDIT_REQUESTS = 'Edit Requests';
    const COLUMN_ASK_REQUESTS = 'Ask Requests';
    const COLUMN_AGENT_REQUESTS = 'Agent Requests';
    const COLUMN_CMDK_USAGES = 'Cmd+K Usages';
    const COLUMN_SUBSCRIPTION_INCLUDED_REQS = 'Subscription Included Reqs';
    const COLUMN_API_KEY_REQS = 'API Key Reqs';
    const COLUMN_USAGE_BASED_REQS = 'Usage Based Reqs';
    const COLUMN_BUGBOT_USAGES = 'Bugbot Usages';
    const COLUMN_MOST_USED_MODEL = 'Most Used Model';
    const COLUMN_MOST_USED_APPLY_EXTENSION = 'Most Used Apply Extension';
    const COLUMN_MOST_USED_TAB_EXTENSION = 'Most Used Tab Extension';
    const COLUMN_CLIENT_VERSION = 'Client Version';

    const SUM_COLUMNS_NUMERIC = [
        COLUMN_CHAT_SUGGESTED_LINES_ADDED, COLUMN_CHAT_SUGGESTED_LINES_DELETED,
        COLUMN_CHAT_ACCEPTED_LINES_ADDED, COLUMN_CHAT_ACCEPTED_LINES_DELETED,
        COLUMN_CHAT_TOTAL_APPLIES, COLUMN_CHAT_TOTAL_ACCEPTS, COLUMN_CHAT_TOTAL_REJECTS,
        COLUMN_CHAT_TABS_SHOWN, COLUMN_TABS_ACCEPTED, COLUMN_EDIT_REQUESTS, COLUMN_ASK_REQUESTS,
        COLUMN_AGENT_REQUESTS, COLUMN_CMDK_USAGES, COLUMN_SUBSCRIPTION_INCLUDED_REQS,
        COLUMN_API_KEY_REQS, COLUMN_USAGE_BASED_REQS, COLUMN_BUGBOT_USAGES
    ];

    // Chart.js Global Defaults (Optional: Customize further as needed)
    // Placed here as they are global settings for Chart.js and used by dashboard functions
    if (typeof Chart !== 'undefined') {
        Chart.defaults.font.family = "'Gilroy', 'Circular', sans-serif";
        Chart.defaults.plugins.legend.position = 'bottom';
        Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(0, 0, 0, 0.7)';
        Chart.defaults.plugins.tooltip.titleFont = { weight: 'bold' };
        Chart.defaults.responsive = true;
        Chart.defaults.maintainAspectRatio = false;
    } else {
        console.warn('Chart.js library not found. Charts will not be rendered.');
    }

    const csvFileInput = document.getElementById('csvFile');
    const loadingDiv = document.getElementById('loading');
    const processMessageEl = document.getElementById('processMessage');
    const downloadAreaEl = document.getElementById('downloadArea');
    const dashboardAreaEl = document.getElementById('dashboardArea');
    const keyMetricsDisplayEl = document.getElementById('keyMetricsDisplay');
    const chartsDisplayEl = document.getElementById('chartsDisplay');
    const reportDatesDivEl = document.getElementById('reportDatesDisplay');

    // --- New Search Functionality Elements ---
    const userSearchInputEl = document.getElementById('userSearchInput');
    const autocompleteSuggestionsEl = document.getElementById('autocompleteSuggestions');
    const searchContainerEl = document.querySelector('.search-container');
    let originalAggregatedData = []; // To store the full dataset
    let currentFilteredData = []; // To store the currently displayed dataset (full or filtered)
    let rawCsvData = []; // Add this at the top level of the IIFE

    if (!csvFileInput) {
        console.error("CRITICAL: CSV File Input element ('csvFile') not found. App cannot initialize.");
        if(processMessageEl) processMessageEl.textContent = "Application initialization error. Required elements missing.";
        return; // Stop script execution
    }

    csvFileInput.addEventListener('change', handleFileSelect);

    // --- Search Event Listeners ---
    if (userSearchInputEl && autocompleteSuggestionsEl) {
        userSearchInputEl.addEventListener('input', handleUserSearchInput);
        userSearchInputEl.addEventListener('focus', handleUserSearchInput); // Show suggestions on focus if input has text
        autocompleteSuggestionsEl.addEventListener('click', handleSuggestionClick);
        document.addEventListener('click', function(event) { // Click outside to hide suggestions
            if (!userSearchInputEl.contains(event.target) && !autocompleteSuggestionsEl.contains(event.target)) {
                autocompleteSuggestionsEl.style.display = 'none';
            }
        });
        userSearchInputEl.addEventListener('keydown', function(event) {
            if (event.key === 'Escape') {
                autocompleteSuggestionsEl.style.display = 'none';
            }
        });
    } else {
        console.warn('User search input or suggestions container not found. Search functionality will be disabled.');
    }

    // --- Drag and Drop Functionality ---
    const fileUploadArea = document.querySelector('.file-upload-area');

    if (fileUploadArea) {
        fileUploadArea.addEventListener('dragover', handleDragOver);
        fileUploadArea.addEventListener('dragleave', handleDragLeave);
        fileUploadArea.addEventListener('drop', handleDrop);

        // Optional: Visual feedback when dragging over the label itself
        const csvFileLabel = document.getElementById('csvFileLabel');
        if (csvFileLabel) {
            fileUploadArea.addEventListener('dragenter', () => csvFileLabel.classList.add('drag-over-label'));
            fileUploadArea.addEventListener('dragleave', () => csvFileLabel.classList.remove('drag-over-label'));
            fileUploadArea.addEventListener('drop', () => csvFileLabel.classList.remove('drag-over-label'));
        }
    } else {
        console.warn("File upload area element ('.file-upload-area') not found. Drag and drop will not work.");
    }

    /**
     * Handles the dragover event on the file upload area.
     * Prevents default behavior to allow drop and adds a visual cue.
     * @param {DragEvent} event - The dragover event.
     */
    function handleDragOver(event) {
        event.preventDefault(); // Necessary to allow dropping
        event.stopPropagation();
        if (fileUploadArea) {
            fileUploadArea.style.borderColor = '#FF00A8'; // Lemonade Pink feedback
        }
    }

    /**
     * Handles the dragleave event on the file upload area.
     * Removes the visual cue.
     * @param {DragEvent} event - The dragleave event.
     */
    function handleDragLeave(event) {
        event.preventDefault();
        event.stopPropagation();
        if (fileUploadArea) {
            fileUploadArea.style.borderColor = '#ddd'; // Reset to default border
        }
    }

    /**
     * Handles the drop event on the file upload area.
     * Processes the dropped file.
     * @param {DragEvent} event - The drop event.
     */
    function handleDrop(event) {
        event.preventDefault();
        event.stopPropagation();
        if (fileUploadArea) {
            fileUploadArea.style.borderColor = '#ddd'; // Reset border
        }

        const files = event.dataTransfer.files;
        if (files.length > 0) {
            // Check if the dropped file is a CSV
            const droppedFile = files[0];
            if (droppedFile.name.endsWith('.csv') || droppedFile.type === 'text/csv' || droppedFile.type === 'application/vnd.ms-excel') {
                csvFileInput.files = files; // Assign to the hidden file input
                const changeEvent = new Event('change', { bubbles: true }); // Create a new change event
                csvFileInput.dispatchEvent(changeEvent); // Dispatch it to trigger handleFileSelect
            } else {
                resetUI('Invalid file type. Please upload a CSV file.');
                console.warn('Invalid file type dropped:', droppedFile.name, droppedFile.type);
            }
        }
    }

    /**
     * Handles the file selection event from the CSV file input.
     * @param {Event} event - The file input change event.
     */
    function handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) {
            resetUI();
            return;
        }
        processCsvFile(file);
    }

    function resetUI(errorMessage = '') {
        if (loadingDiv) loadingDiv.style.display = 'none';
        if (processMessageEl) {
            processMessageEl.textContent = errorMessage;
            processMessageEl.style.color = errorMessage ? 'red' : 'inherit';
        }
        if (downloadAreaEl) downloadAreaEl.innerHTML = '';
        if (dashboardAreaEl) {
            dashboardAreaEl.style.display = 'none';
            dashboardAreaEl.innerHTML = ''; // Also clear its content if reset implies full clear
        }
        if (keyMetricsDisplayEl) keyMetricsDisplayEl.innerHTML = '';
        if (chartsDisplayEl) chartsDisplayEl.innerHTML = '';
        if (reportDatesDivEl) reportDatesDivEl.style.display = 'none';
        if (searchContainerEl) searchContainerEl.style.display = 'none';
    }

    function displayLoadingState() {
        if (loadingDiv) loadingDiv.style.display = 'block';
        if (processMessageEl) processMessageEl.textContent = ''; // Clear previous messages
        if (downloadAreaEl) downloadAreaEl.innerHTML = ''; // Clear previous buttons
        if (dashboardAreaEl) dashboardAreaEl.style.display = 'none'; // Hide dashboard while processing
        if (keyMetricsDisplayEl) keyMetricsDisplayEl.innerHTML = ''; // Clear previous metrics
        if (chartsDisplayEl) chartsDisplayEl.innerHTML = ''; // Clear previous charts
        if (reportDatesDivEl) reportDatesDivEl.style.display = 'none'; // Hide report dates
        if (searchContainerEl) searchContainerEl.style.display = 'none';
    }

    function processCsvFile(file) {
        displayLoadingState();

        Papa.parse(file, {
            header: true,
            dynamicTyping: false,
            skipEmptyLines: true,
            complete: function(results) {
                if (loadingDiv) loadingDiv.style.display = 'none';

                console.log('[DEBUG] PapaParse complete. results:', results);

                if (results.errors.length > 0) {
                    console.error("[DEBUG] Parsing errors found:", results.errors);
                    const errorMessages = results.errors.map(err => err.message).join(', ');
                    resetUI(`Error parsing CSV: ${errorMessages}. Check console for details.`);
                    return;
                }
                console.log('[DEBUG] No parsing errors.');

                const data = results.data;
                console.log('[DEBUG] Parsed data (results.data):', data);

                if (!data || data.length === 0) {
                    console.warn('[DEBUG] No data rows found in CSV.');
                    resetUI('No data found in the CSV.');
                    return;
                }
                console.log(`[DEBUG] ${data.length} data rows found.`);

                try {
                    console.log('[DEBUG] Calling processParsedData...');
                    processParsedData(data);
                } catch (error) {
                    console.error('[DEBUG] Error during data processing:', error);
                    resetUI(`Error during data processing: ${error.message}. Check console.`);
                }
            },
            error: function(err, file, inputElem, reason) {
                console.error("CSV parsing error object:", err);
                console.error("CSV parsing error reason:", reason);
                resetUI(`Error reading CSV: ${reason || 'Unknown error'}. Check console.`);
            }
        });
    }

    /**
     * Processes the data parsed from the CSV file, aggregates it, and updates the UI.
     * @param {Array<Object<string, string>>} data - The array of data rows parsed from the CSV.
     */
    function processParsedData(data) {
        rawCsvData = data; // Store the raw data
        console.log('[DEBUG] Calling aggregateCsvData...');
        const aggregationResult = aggregateCsvData(data);
        console.log('[DEBUG] aggregateCsvData returned:', aggregationResult);

        const { aggregatedData, filename, reportStartDate, reportEndDate } = aggregationResult;

        if (!aggregatedData || aggregatedData.length === 0) {
            console.warn('[DEBUG] No data could be aggregated.');
            resetUI('No data could be aggregated to display a dashboard or generate a report.');
            return;
        }

        originalAggregatedData = JSON.parse(JSON.stringify(aggregatedData)); // Store a deep copy
        currentFilteredData = [...originalAggregatedData]; // Initially, display all data

        console.log('[DEBUG] originalAggregatedData stored, currentFilteredData initialized.');
        console.log('[DEBUG] currentFilteredData has data, calling displayDashboardMetrics...');

        if (chartsDisplayEl) {
            displayDashboardMetrics(currentFilteredData, chartsDisplayEl); // Use currentFilteredData
        } else {
            console.error("Error: chartsDisplay element not found before calling displayDashboardMetrics.");
            // Optionally, inform the user or log this more visibly
        }


        displayReportDates(reportStartDate, reportEndDate);
        createAndOfferDownload(aggregatedData, filename); // Download link should use the original full data

        if (processMessageEl) {
            processMessageEl.textContent = 'CSV processed. Dashboard generated and report ready for download.';
            processMessageEl.style.color = 'green';
        }
        if (dashboardAreaEl) dashboardAreaEl.style.display = 'block'; // Show the dashboard area
        if (searchContainerEl) searchContainerEl.style.display = 'block';
    }

    function displayReportDates(startDate, endDate) {
        if (reportDatesDivEl) {
            let dateText = 'Report Dates: Not Available';
            if (startDate !== 'NODATE' && endDate !== 'NODATE') {
                dateText = `Report Period: ${startDate} to ${endDate}`;
            } else if (startDate !== 'NODATE') {
                dateText = `Report Date: ${startDate}`;
            } else if (endDate !== 'NODATE') {
                dateText = `Report Date: ${endDate}`;
            }
            reportDatesDivEl.textContent = dateText;
            reportDatesDivEl.style.display = 'block';
        }
    }

    function createAndOfferDownload(dataForCsv, filename) {
        if (downloadAreaEl) {
            const csv = Papa.unparse(dataForCsv);
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);

            const downloadLink = document.createElement('a');
            downloadLink.setAttribute('href', url);
            downloadLink.setAttribute('download', filename);
            downloadLink.textContent = 'Download Aggregated CSV';
            downloadLink.id = 'downloadLink'; // Maintain ID for styling

            downloadAreaEl.innerHTML = ''; // Clear previous links
            downloadAreaEl.appendChild(downloadLink);
        }
    }

    /**
     * Aggregates CSV data by email, processing various metrics.
     * @param {Array<Object<string, string>>} data - The array of data rows parsed from the CSV.
     * @returns {{aggregatedData: Array<Object>, filename: string, reportStartDate: string, reportEndDate: string}} An object containing the aggregated data, the generated filename, and the report start/end dates.
     */
    function aggregateCsvData(data) {
        const emailMap = new Map();

        // Calculate overall date range for the report first
        const { overallOldestDate, overallNewestDate } = calculateOverallDateRange(data);

        data.forEach(row => {
            const email = row[COLUMN_EMAIL]?.trim();
            if (!email) {
                return; // Skip rows without an email address
            }

            let entry = emailMap.get(email);
            if (!entry) {
                entry = initializeEmailEntry(row, email);
                emailMap.set(email, entry);
            }
            updateEmailEntry(entry, row);
        });

        const formattedOverallOldestDate = overallOldestDate ? formatDateWithoutTime(overallOldestDate) : 'NODATE';
        const formattedOverallNewestDate = overallNewestDate ? formatDateWithoutTime(overallNewestDate) : 'NODATE';
        const generatedFilename = generateReportFilename(formattedOverallOldestDate, formattedOverallNewestDate);

        const finalResults = Array.from(emailMap.values()).map(formatFinalEntry);

        console.log('[DEBUG] Final aggregated results for CSV:', finalResults);

        return {
            aggregatedData: finalResults,
            filename: generatedFilename,
            reportStartDate: formattedOverallOldestDate,
            reportEndDate: formattedOverallNewestDate
        };
    }

    /**
     * Calculates the overall start and end dates from all row entries.
     * @param {Array<Object<string, string>>} data - The raw data rows from the CSV.
     * @returns {{overallOldestDate: Date | null, overallNewestDate: Date | null}} An object containing the earliest and latest dates found.
     */
    function calculateOverallDateRange(data) {
        let overallOldestDate = null;
        let overallNewestDate = null;

        data.forEach(row => {
            const currentDateObj = parseDate(row[COLUMN_DATE]);
            if (currentDateObj) {
                if (!overallOldestDate || currentDateObj < overallOldestDate) {
                    overallOldestDate = currentDateObj;
                }
                if (!overallNewestDate || currentDateObj > overallNewestDate) {
                    overallNewestDate = currentDateObj;
                }
            }
        });
        return { overallOldestDate, overallNewestDate };
    }

    /**
     * Generates a filename for the report based on a date range.
     * @param {string} startDateStr - Formatted start date string (e.g., "DD-MM-YYYY" or "NODATE").
     * @param {string} endDateStr - Formatted end date string (e.g., "DD-MM-YYYY" or "NODATE").
     * @returns {string} The generated filename.
     */
    function generateReportFilename(startDateStr, endDateStr) {
        if (startDateStr === 'NODATE' && endDateStr === 'NODATE') {
            return `cursor_aggregated_data_unknown_period.csv`;
        }
        if (startDateStr === 'NODATE') {
             return `cursor_aggregated_data_ending_${endDateStr}.csv`;
        }
        if (endDateStr === 'NODATE') {
            return `cursor_aggregated_data_starting_${startDateStr}.csv`;
        }
        return `cursor_aggregated_data_${startDateStr}_to_${endDateStr}.csv`;
    }

    /**
     * Initializes a new entry object for the emailMap, used in data aggregation.
     * @param {Object<string, string>} row - The current CSV row being processed (used for User ID).
     * @param {string} email - The email address for the entry.
     * @returns {Object} The initialized entry object with default values for aggregation.
     */
    function initializeEmailEntry(row, email) {
        return {
            [COLUMN_EMAIL]: email,
            oldestDateObj: null,
            newestDateObj: null,
            [COLUMN_USER_ID]: row[COLUMN_USER_ID] || '',
            [COLUMN_IS_ACTIVE]: false,
            [COLUMN_CHAT_SUGGESTED_LINES_ADDED]: 0,
            [COLUMN_CHAT_SUGGESTED_LINES_DELETED]: 0,
            [COLUMN_CHAT_ACCEPTED_LINES_ADDED]: 0,
            [COLUMN_CHAT_ACCEPTED_LINES_DELETED]: 0,
            [COLUMN_CHAT_TOTAL_APPLIES]: 0,
            [COLUMN_CHAT_TOTAL_ACCEPTS]: 0,
            [COLUMN_CHAT_TOTAL_REJECTS]: 0,
            [COLUMN_CHAT_TABS_SHOWN]: 0,
            [COLUMN_TABS_ACCEPTED]: 0,
            [COLUMN_EDIT_REQUESTS]: 0,
            [COLUMN_ASK_REQUESTS]: 0,
            [COLUMN_AGENT_REQUESTS]: 0,
            [COLUMN_CMDK_USAGES]: 0,
            [COLUMN_SUBSCRIPTION_INCLUDED_REQS]: 0,
            [COLUMN_API_KEY_REQS]: 0,
            [COLUMN_USAGE_BASED_REQS]: 0,
            [COLUMN_BUGBOT_USAGES]: 0,
            modelCounts: new Map(),
            applyExtensionCounts: new Map(),
            tabExtensionCounts: new Map(),
            [COLUMN_CLIENT_VERSION]: '',
            userDates: []
        };
    }

    /**
     * Updates an existing entry in the emailMap with data from the current CSV row.
     * Modifies the entry object directly.
     * @param {Object} entry - The existing entry for the email.
     * @param {Object<string, string>} row - The current CSV row being processed.
     */
    function updateEmailEntry(entry, row) {
        // Date - store all, find min/max for user later
        const currentDateObj = parseDate(row[COLUMN_DATE]);
        if (currentDateObj) {
            entry.userDates.push(currentDateObj);
        }

        // Is Active - boolean OR
        const isActive = (row[COLUMN_IS_ACTIVE] || '').toUpperCase() === 'TRUE';
        entry[COLUMN_IS_ACTIVE] = entry[COLUMN_IS_ACTIVE] || isActive;

        // Sums
        SUM_COLUMNS_NUMERIC.forEach(col => {
            const value = parseFloat(row[col]);
            if (!isNaN(value)) {
                entry[col] = (entry[col] || 0) + value;
            }
        });

        // Most Used
        const model = row[COLUMN_MOST_USED_MODEL]?.trim();
        if (model) {
            entry.modelCounts.set(model, (entry.modelCounts.get(model) || 0) + 1);
        }
        const applyExt = row[COLUMN_MOST_USED_APPLY_EXTENSION]?.trim();
        if (applyExt) {
            entry.applyExtensionCounts.set(applyExt, (entry.applyExtensionCounts.get(applyExt) || 0) + 1);
        }
        const tabExt = row[COLUMN_MOST_USED_TAB_EXTENSION]?.trim();
        if (tabExt) {
            entry.tabExtensionCounts.set(tabExt, (entry.tabExtensionCounts.get(tabExt) || 0) + 1);
        }

        // Client Version - highest
        const currentVersion = row[COLUMN_CLIENT_VERSION]?.trim();
        if (currentVersion) {
            if (!entry[COLUMN_CLIENT_VERSION] || compareVersions(currentVersion, entry[COLUMN_CLIENT_VERSION]) > 0) {
                entry[COLUMN_CLIENT_VERSION] = currentVersion;
            }
        }
    }

    /**
     * Formats a raw aggregated entry into its final structure for the CSV output and dashboard display.
     * @param {Object} rawEntry - The raw aggregated data for a single user, including intermediate data like `userDates` and `modelCounts`.
     * @returns {Object} The formatted entry with user-friendly date strings and calculated "most used" fields.
     */
    function formatFinalEntry(rawEntry) {
        if (rawEntry.userDates.length > 0) {
            rawEntry.oldestDateObj = new Date(Math.min(...rawEntry.userDates));
            rawEntry.newestDateObj = new Date(Math.max(...rawEntry.userDates));
        }

        const oldestDateStr = rawEntry.oldestDateObj ? formatDateWithoutTime(rawEntry.oldestDateObj) : '';
        const newestDateStr = rawEntry.newestDateObj ? formatDateWithoutTime(rawEntry.newestDateObj) : '';
        const datesValue = (oldestDateStr && newestDateStr) ? `${oldestDateStr} / ${newestDateStr}` : (oldestDateStr || newestDateStr || '');

        const mostUsedModel = getMostFrequent(rawEntry.modelCounts);
        const mostUsedApplyExt = getMostFrequent(rawEntry.applyExtensionCounts);
        const mostUsedTabExt = getMostFrequent(rawEntry.tabExtensionCounts);
        const isActiveStr = rawEntry[COLUMN_IS_ACTIVE] ? 'TRUE' : 'FALSE';

        return {
            'Dates': datesValue,
            [COLUMN_EMAIL]: rawEntry[COLUMN_EMAIL],
            [COLUMN_USER_ID]: rawEntry[COLUMN_USER_ID],
            [COLUMN_IS_ACTIVE]: isActiveStr,
            [COLUMN_CHAT_SUGGESTED_LINES_ADDED]: rawEntry[COLUMN_CHAT_SUGGESTED_LINES_ADDED],
            [COLUMN_CHAT_SUGGESTED_LINES_DELETED]: rawEntry[COLUMN_CHAT_SUGGESTED_LINES_DELETED],
            [COLUMN_CHAT_ACCEPTED_LINES_ADDED]: rawEntry[COLUMN_CHAT_ACCEPTED_LINES_ADDED],
            [COLUMN_CHAT_ACCEPTED_LINES_DELETED]: rawEntry[COLUMN_CHAT_ACCEPTED_LINES_DELETED],
            [COLUMN_CHAT_TOTAL_APPLIES]: rawEntry[COLUMN_CHAT_TOTAL_APPLIES],
            [COLUMN_CHAT_TOTAL_ACCEPTS]: rawEntry[COLUMN_CHAT_TOTAL_ACCEPTS],
            [COLUMN_CHAT_TOTAL_REJECTS]: rawEntry[COLUMN_CHAT_TOTAL_REJECTS],
            [COLUMN_CHAT_TABS_SHOWN]: rawEntry[COLUMN_CHAT_TABS_SHOWN],
            [COLUMN_TABS_ACCEPTED]: rawEntry[COLUMN_TABS_ACCEPTED],
            [COLUMN_EDIT_REQUESTS]: rawEntry[COLUMN_EDIT_REQUESTS],
            [COLUMN_ASK_REQUESTS]: rawEntry[COLUMN_ASK_REQUESTS],
            [COLUMN_AGENT_REQUESTS]: rawEntry[COLUMN_AGENT_REQUESTS],
            [COLUMN_CMDK_USAGES]: rawEntry[COLUMN_CMDK_USAGES],
            [COLUMN_SUBSCRIPTION_INCLUDED_REQS]: rawEntry[COLUMN_SUBSCRIPTION_INCLUDED_REQS],
            [COLUMN_API_KEY_REQS]: rawEntry[COLUMN_API_KEY_REQS],
            [COLUMN_USAGE_BASED_REQS]: rawEntry[COLUMN_USAGE_BASED_REQS],
            [COLUMN_BUGBOT_USAGES]: rawEntry[COLUMN_BUGBOT_USAGES],
            [COLUMN_MOST_USED_MODEL]: mostUsedModel,
            [COLUMN_MOST_USED_APPLY_EXTENSION]: mostUsedApplyExt,
            [COLUMN_MOST_USED_TAB_EXTENSION]: mostUsedTabExt,
            [COLUMN_CLIENT_VERSION]: rawEntry[COLUMN_CLIENT_VERSION]
        };
    }

    function parseDate(dateString) {
        if (!dateString) return null;

        // Try ISO format first (most common in our data)
        const isoDate = new Date(dateString);
        if (!isNaN(isoDate.getTime())) {
            return isoDate;
        }

        // Fallback to DD-MM-YYYY format for backward compatibility
        const parts = dateString.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
        if (parts) {
            const date = new Date(parts[3], parts[2] - 1, parts[1]); // Month is 0-indexed
            if (!isNaN(date.getTime())) {
                return date;
            }
        }

        console.warn("Could not parse date:", dateString);
        return null;
    }

    function formatDateWithoutTime(date) {
        if (!date) return '';
        const year = date.getFullYear();
        const month = (date.getMonth() + 1).toString().padStart(2, '0'); // Month is 0-indexed
        const day = date.getDate().toString().padStart(2, '0');
        return `${year}-${month}-${day}`; // Return the formatted date string
    }

    function compareVersions(v1, v2) {
        if (v1 === v2) return 0;
        if (!v1) return -1; // Consider empty/null v1 as "lesser"
        if (!v2) return 1;  // Consider empty/null v2 as "lesser"

        const parts1 = v1.split('.').map(Number);
        const parts2 = v2.split('.').map(Number);
        const maxLength = Math.max(parts1.length, parts2.length);

        for (let i = 0; i < maxLength; i++) {
            const p1 = parts1[i] || 0;
            const p2 = parts2[i] || 0;
            if (p1 > p2) return 1;
            if (p1 < p2) return -1;
        }
        return 0;
    }

    function getMostFrequent(countsMap) {
        if (!countsMap || countsMap.size === 0) {
            return '';
        }

        let maxCount = 0;
        countsMap.forEach(count => {
            if (count > maxCount) {
                maxCount = count;
            }
        });

        const mostFrequentItems = [];
        countsMap.forEach((count, item) => {
            if (count === maxCount) {
                mostFrequentItems.push(item);
            }
        });

        mostFrequentItems.sort(); // Sort for consistent output

        return mostFrequentItems.join(', ');
    }

    function formatPercentage(value) {
        if (value === null || typeof value === 'undefined' || isNaN(value)) return 'N/A';
        return (value * 100).toFixed(1) + '%';
    }

    /**
     * Creates a chart canvas element and appends it to a container.
     * @param {string} title - The title of the chart.
     * @param {string} chartId - The ID for the canvas element.
     * @param {HTMLElement} containerElement - The HTML element to append the chart to.
     * @returns {CanvasRenderingContext2D | null} The 2D rendering context of the canvas, or null if creation fails.
     */
    function createChartCanvas(title, chartId, containerElement) {
        if (!containerElement) {
            console.error(`Error: Container element for chart '${title}' (ID: ${chartId}) not found.`);
            return null;
        }
        // Clear previous content if any (e.g. error messages)
        // containerElement.innerHTML = ''; // This might be too aggressive if other elements are siblings

        const chartWrapper = document.createElement('div');
        chartWrapper.className = 'chart-container'; // Use existing class for styling

        const titleEl = document.createElement('h4');
        titleEl.textContent = title;
        // Styles for h4 are in CSS, no need to set them here unless overriding
        chartWrapper.appendChild(titleEl);

        const canvas = document.createElement('canvas');
        canvas.id = chartId;
        // Height/width for canvas is managed by Chart.js options and CSS for chart-container
        chartWrapper.appendChild(canvas);

        containerElement.appendChild(chartWrapper);
        return canvas.getContext('2d');
    }

    /**
     * Renders a key-value metric display.
     * @param {string} title - The title of the metric.
     * @param {string|number} value - The value of the metric.
     * @param {string} [unit=''] - The unit for the metric value.
     * @param {string} [definition=''] - An optional definition or description for the metric.
     * @param {HTMLElement} containerElement - The HTML element to append the metric display to.
     */
    function renderKeyValueMetric(title, value, unit = '', definition = '', containerElement) {
        if (!containerElement) {
            console.error(`Error: Container element for key metric '${title}' not found.`);
            return;
        }

        // Use the structure defined in index.html for .metric-card
        const metricCard = document.createElement('div');
        metricCard.className = 'metric-card';

        const titleEl = document.createElement('h4');
        titleEl.textContent = title;
        metricCard.appendChild(titleEl);

        const valueEl = document.createElement('div');
        valueEl.className = 'metric-value';

        let displayValue = value;
        if (typeof value === 'number') {
            if (value % 1 !== 0) { // Check if it's a float
                displayValue = value.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 });
            } else {
                displayValue = value.toLocaleString();
            }
        }
        valueEl.textContent = displayValue;
        metricCard.appendChild(valueEl);

        if (unit) {
            const unitEl = document.createElement('div');
            unitEl.className = 'metric-unit';
            unitEl.textContent = unit;
            metricCard.appendChild(unitEl);
        }

        // If definition is needed, it can be added as a tooltip or a small paragraph
        // For now, relying on the HTML structure which doesn't explicitly show definition easily

        containerElement.appendChild(metricCard);
    }

    /**
     * Renders a doughnut chart for developer status (active vs. inactive).
     * @param {number} activeCount - Number of active developers.
     * @param {number} inactiveCount - Number of inactive/churned developers.
     * @param {HTMLElement} containerElement - The HTML element (typically a div with class 'charts-grid') to append the chart's container to.
     */
    function renderDevStatusChart(activeCount, inactiveCount, containerElement) {
        const ctx = createChartCanvas('Developer Status (Active vs. Inactive)', 'devStatusChart', containerElement);
        if (!ctx) return;

        new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Active Developers', 'Inactive/Churned Developers'],
                datasets: [{
                    label: 'Developer Status',
                    data: [activeCount, inactiveCount],
                    backgroundColor: [
                        '#FF00A8', // Lemonade Pink
                        '#7C4DFF'  // Soft Purple
                    ],
                    borderColor: [
                        '#FFFFFF',
                        '#FFFFFF'
                    ],
                    borderWidth: 2
                }]
            },
            options: {
                plugins: {
                    title: {
                        display: false, // Title is handled by createChartCanvas
                        text: 'Developer Status'
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                let label = context.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed !== null) {
                                    label += context.parsed.toLocaleString();
                                }
                                return label;
                            }
                        }
                    }
                }
            }
        });
    }

    /**
     * Renders a single value horizontal bar chart, often used for displaying a specific metric value quantitatively.
     * @param {string} chartTitle - The title to display above the chart.
     * @param {string} dataLabel - The label for the data point itself (e.g., the name of the metric being shown on the bar).
     * @param {number} value - The numerical value of the metric.
     * @param {string} [unit=''] - The unit for the value (e.g., "lines", "ms", "usages"). Appended to tooltip values.
     * @param {string} [definition=''] - An optional definition or further explanation for the metric, often shown in the tooltip footer.
     * @param {string} chartId - The unique ID to be assigned to the canvas element for this chart.
     * @param {HTMLElement} containerElement - The HTML element (typically a div with class 'charts-grid') to append the chart's container to.
     */
    function renderSingleValueBarChart(chartTitle, dataLabel, value, unit = '', definition = '', chartId, containerElement) {
        const ctx = createChartCanvas(chartTitle, chartId, containerElement);
        if (!ctx) return;

        // Ensure value is a number, default to 0 if not
        const numericValue = typeof value === 'number' && !isNaN(value) ? value : 0;

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: [dataLabel], // The item being measured
                datasets: [{
                    label: unit || dataLabel, // Use unit if available, else dataLabel
                    data: [numericValue],
                    backgroundColor: ['#00C4FF'], // Lemonade Blue
                    borderColor: ['#00A0D1'],
                    borderWidth: 1,
                    barPercentage: 0.5, // Makes the bar narrower
                    categoryPercentage: 0.8
                }]
            },
            options: {
                indexAxis: 'y', // Horizontal bar for single value often looks good
                scales: {
                    x: {
                        beginAtZero: true,
                        ticks: {
                            precision: 0 // Show whole numbers if applicable
                        }
                    },
                    y: {
                       display: false // Hide y-axis label for single bar
                    }
                },
                plugins: {
                    legend: {
                        display: false // Usually not needed for a single bar
                    },
                    title: {
                        display: false,
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed.x !== null) {
                                    label += context.parsed.x.toLocaleString() + (unit ? ' ' + unit : '');
                                }
                                return label;
                            },
                            footer: function() {
                                return definition ? definition : '';
                            }
                        }
                    }
                }
            }
        });
    }

    /**
     * Renders a percentage gauge chart (half doughnut) to visually represent a value from 0% to 100%.
     * @param {string} chartTitle - The title to display above the chart.
     * @param {number} value - The percentage value, expected to be between 0.0 (0%) and 1.0 (100%).
     * @param {string} [definition=''] - An optional definition for the metric, which can be shown in the tooltip footer.
     * @param {string} chartId - The unique ID to be assigned to the canvas element for this chart.
     * @param {HTMLElement} containerElement - The HTML element (typically a div with class 'charts-grid') to append the chart's container to.
     */
    function renderPercentageGaugeChart(chartTitle, value, definition = '', chartId, containerElement) {
        const ctx = createChartCanvas(chartTitle, chartId, containerElement);
        if (!ctx) return;

        // Ensure value is a number between 0 and 1, default to 0 if not
        const normalizedValue = (typeof value === 'number' && !isNaN(value) && value >= 0 && value <= 1) ? value : 0;
        const percentage = parseFloat((normalizedValue * 100).toFixed(1));

        new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['Achieved', 'Remaining'],
                datasets: [{
                    data: [percentage, 100 - percentage],
                    backgroundColor: ['#FF00A8', '#E0E0E0'], // Pink for achieved, Grey for remaining
                    borderWidth: 0,
                    circumference: 180, // Half circle
                    rotation: 270      // Start from bottom
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true, // Let Chart.js handle aspect ratio
                cutout: '70%', // Makes it a gauge
                plugins: {
                    legend: {
                        display: false
                    },
                    title: {
                        display: false,
                    },
                    tooltip: {
                        enabled: true, // Enable tooltips
                        callbacks: {
                            // Keep the default title or provide a simple one if needed
                            // title: function(tooltipItems) { return chartTitle; },
                            label: function(tooltipItem) {
                                // Display the segment label (e.g., "Achieved") and its value
                                let label = tooltipItem.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (tooltipItem.parsed !== null) {
                                    label += parseFloat(tooltipItem.parsed).toFixed(1) + '%';
                                }
                                return label;
                            },
                            footer: function(tooltipItems) {
                                // Add the definition to the footer of the tooltip
                                // The 'definition' variable is accessible from the outer function's scope
                                return definition ? definition : '';
                            }
                        }
                    },
                    // Custom plugin to display text in center
                    centerText: {
                        display: true,
                        text: percentage + '%',
                        color: '#FF00A8', // Pink
                        fontStyle: 'bold',
                        fontSize: Math.min(ctx.canvas.width / 5, ctx.canvas.height / 3) // Dynamic font size
                    }
                }
            },
            plugins: [{ // Registering the custom plugin
                id: 'centerText',
                beforeDraw: (chart) => {
                    const options = chart.options.plugins.centerText;
                    if (options && options.display) {
                        const {ctx, width, height} = chart;
                        ctx.restore();
                        const fontSize = options.fontSize || (height / 114).toFixed(2); // Fallback dynamic size
                        ctx.font = `${options.fontStyle || 'normal'} ${fontSize}px ${Chart.defaults.font.family}`;
                        ctx.textBaseline = 'middle';
                        ctx.textAlign = 'center';
                        ctx.fillStyle = options.color || Chart.defaults.color;
                        
                        // Position in the middle of the doughnut hole
                        const x = width / 2;
                        const y = height * 0.75; // Adjust for half doughnut

                        ctx.fillText(options.text, x, y);
                        ctx.save();
                    }
                }
            }]
        });
    }

    /**
     * Renders a bar chart showing the popularity of different models based on active user usage.
     * Each bar represents a model, and its length corresponds to the percentage of active users who used that model.
     * @param {Array<{model: string, percentage: number}>} popularityData - An array of objects, where each object contains a `model` name (string) and its usage `percentage` (number).
     * @param {HTMLElement} containerElement - The HTML element (typically a div with class 'charts-grid') to append the chart's container to.
     */
    function renderModelPopularityChart(popularityData, containerElement) {
        const ctx = createChartCanvas('Model Popularity (% Active Users)', 'modelPopularityChart', containerElement);
        if (!ctx) return;

        if (!popularityData || popularityData.length === 0) {
            // Display a message directly within the chart's designated space
            const chartWrapper = ctx.canvas.parentElement; // This should be the .chart-container div
            if (chartWrapper) {
                chartWrapper.innerHTML = '<h4>Model Popularity (% Active Users)</h4><p style="text-align: center; padding-top: 20px;">No model popularity data available.</p>';
            } else {
                // Fallback if the parent isn't found as expected, though less ideal
                containerElement.innerHTML += '<div class="chart-container"><h4>Model Popularity (% Active Users)</h4><p>No model popularity data available.</p></div>';
            }
            return;
        }

        const labels = popularityData.map(item => item.model);
        const data = popularityData.map(item => item.percentage);

        new Chart(ctx, {
            type: 'bar', // Horizontal bar chart is good for this
            data: {
                labels: labels,
                datasets: [{
                    label: '% of Active Users',
                    data: data,
                    backgroundColor: '#391085', // Dark Purple/Blue
                    borderColor: '#2A0A64',
                    borderWidth: 1
                }]
            },
            options: {
                indexAxis: 'y', // Makes it horizontal
                scales: {
                    x: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return value + '%';
                            }
                        }
                    }
                },
                plugins: {
                    title: {
                        display: false,
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                return context.dataset.label + ': ' + context.parsed.x.toFixed(1) + '%';
                            }
                        }
                    }
                }
            }
        });
    }

    /**
     * Renders a bar chart showing the top N power users based on a specific metric (currently AI lines added).
     * Each bar represents a user (email prefix), and its length corresponds to their metric value.
     * @param {Array<{email: string, aiLines: number}>} usersData - An array of user objects, each containing `email` (string) and `aiLines` (number).
     * @param {HTMLElement} containerElement - The HTML element (typically a div with class 'charts-grid') to append the chart's container to.
     */
    function renderTopPowerUsersChart(usersData, containerElement) {
        const ctx = createChartCanvas('Top 10 Power Users (AI Lines Added)', 'topPowerUsersChart', containerElement);
        if (!ctx) return;

        if (!usersData || usersData.length === 0) {
             // Display a message directly within the chart's designated space
            const chartWrapper = ctx.canvas.parentElement; // This should be the .chart-container div
            if (chartWrapper) {
                chartWrapper.innerHTML = '<h4>Top 10 Power Users (AI Lines Added)</h4><p style="text-align: center; padding-top: 20px;">No power user data available.</p>';
            } else {
                containerElement.innerHTML += '<div class="chart-container"><h4>Top 10 Power Users (AI Lines Added)</h4><p>No power user data available.</p></div>';
            }
            return;
        }

        const labels = usersData.map(user => user.email.substring(0, user.email.indexOf('@') > 0 ? user.email.indexOf('@') : 20)); // Shorten email
        const data = usersData.map(user => user.aiLines);

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'AI Lines Added',
                    data: data,
                    backgroundColor: '#FF00A8', // Lemonade Pink
                    borderColor: '#D1008F',
                    borderWidth: 1
                }]
            },
            options: {
                indexAxis: 'y',
                scales: {
                    x: {
                        beginAtZero: true
                    }
                },
                plugins: {
                    title: {
                        display: false,
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                return context.dataset.label + ': ' + context.parsed.x.toLocaleString();
                            }
                        }
                    }
                }
            }
        });
    }

    /**
     * Calculates the number of active developers from the aggregated data.
     * @param {Array<Object>} data - The aggregated data, where each object represents a user.
     * @returns {number} The count of users marked as active.
     */
    function getActiveDeveloperCount(data) {
        return data.filter(row => row[COLUMN_IS_ACTIVE] === 'TRUE').length;
    }

    /**
     * Calculates the number of inactive (or churned) developers from the aggregated data.
     * @param {Array<Object>} data - The aggregated data, where each object represents a user.
     * @returns {number} The count of users marked as inactive.
     */
    function getInactiveDeveloperCount(data) {
        return data.filter(row => row[COLUMN_IS_ACTIVE] === 'FALSE').length;
    }

    /**
     * Retrieves a list of inactive users (email and user ID).
     * @param {Array<Object>} data - The aggregated data.
     * @returns {Array<{email: string, userId: string}>} An array of objects, each containing the email and user ID of an inactive user.
     */
    function getInactiveUsersList(data) {
        return data
            .filter(row => row[COLUMN_IS_ACTIVE] === 'FALSE')
            .map(user => ({
                email: user[COLUMN_EMAIL],
                userId: user[COLUMN_USER_ID]
            }))
            .sort((a,b) => a.email.localeCompare(b.email)); // Sort alphabetically by email
    }

    /**
     * Renders the list of inactive users into the designated UI element.
     * @param {Array<{email: string, userId: string}>} inactiveUsers - An array of inactive user objects.
     */
    function renderInactiveUsersList(inactiveUsers) {
        const inactiveUserListEl = document.getElementById('inactiveUserList');
        const inactiveUsersDisplayEl = document.getElementById('inactiveUsersDisplay');

        if (!inactiveUserListEl || !inactiveUsersDisplayEl) {
            console.warn('Inactive users list element or container not found. Skipping rendering.');
            return;
        }

        inactiveUserListEl.innerHTML = ''; // Clear previous list

        if (inactiveUsers.length === 0) {
            inactiveUserListEl.innerHTML = '<li style="text-align: center; color: #777;">No inactive users found.</li>';
            // Optionally hide the whole section if no inactive users, or adjust styling
            // inactiveUsersDisplayEl.style.display = 'none'; 
        } else {
            inactiveUsers.forEach(user => {
                const listItem = document.createElement('li');
                listItem.textContent = `${user.email} (User ID: ${user.userId || 'N/A'})`;
                listItem.style.padding = '5px 0';
                listItem.style.borderBottom = '1px solid #eee';
                inactiveUserListEl.appendChild(listItem);
            });
            inactiveUsersDisplayEl.style.display = 'block'; // Ensure it's visible
        }
    }

    /**
     * Calculates the total AI-generated lines added (sum of 'Chat Suggested Lines Added' and 'Chat Accepted Lines Added') across all users.
     * @param {Array<Object>} data - The aggregated data.
     * @returns {number} The total number of AI-generated lines added.
     */
    function getTotalAIGeneratedLinesAdded(data) {
        return data.reduce((sum, row) => {
            const suggested = parseFloat(row[COLUMN_CHAT_SUGGESTED_LINES_ADDED]) || 0;
            const accepted = parseFloat(row[COLUMN_CHAT_ACCEPTED_LINES_ADDED]) || 0;
            return sum + suggested + accepted;
        }, 0);
    }

    /**
     * Calculates the overall acceptance rate of AI suggestions (Total Accepts / Total Applies).
     * @param {Array<Object>} data - The aggregated data.
     * @returns {number | null} The acceptance rate as a decimal (0.0 to 1.0), or null if no suggestions were applied (to avoid division by zero).
     */
    function getAcceptanceRateOfAISuggestions(data) {
        const totalAccepts = data.reduce((sum, row) => sum + (parseFloat(row[COLUMN_CHAT_TOTAL_ACCEPTS]) || 0), 0);
        const totalApplies = data.reduce((sum, row) => sum + (parseFloat(row[COLUMN_CHAT_TOTAL_APPLIES]) || 0), 0);
        if (totalApplies === 0) return null; // Avoid division by zero, return null or 0 based on preference
        return totalAccepts / totalApplies;
    }

    /**
     * Calculates the overall conversion rate of tab suggestions (Tabs Accepted / Tabs Shown).
     * @param {Array<Object>} data - The aggregated data.
     * @returns {number | null} The tab conversion rate as a decimal (0.0 to 1.0), or null if no tab suggestions were shown.
     */
    function getTabConversionRate(data) {
        const tabsAccepted = data.reduce((sum, row) => sum + (parseFloat(row[COLUMN_TABS_ACCEPTED]) || 0), 0);
        const chatTabsShown = data.reduce((sum, row) => sum + (parseFloat(row[COLUMN_CHAT_TABS_SHOWN]) || 0), 0);
        if (chatTabsShown === 0) return null; // Avoid division by zero
        return tabsAccepted / chatTabsShown;
    }

    /**
     * Calculates the prompt activity index, defined as the sum of Ask, Edit, and Agent requests across all users.
     * @param {Array<Object>} data - The aggregated data.
     * @returns {number} The total prompt activity index.
     */
    function getPromptActivityIndex(data) {
        return data.reduce((sum, row) => {
            const ask = parseFloat(row[COLUMN_ASK_REQUESTS]) || 0;
            const edit = parseFloat(row[COLUMN_EDIT_REQUESTS]) || 0;
            const agent = parseFloat(row[COLUMN_AGENT_REQUESTS]) || 0;
            return sum + ask + edit + agent;
        }, 0);
    }

    /**
     * Calculates the total number of Cmd+K feature usages across all users.
     * @param {Array<Object>} data - The aggregated data.
     * @returns {number} The total Cmd+K usages.
     */
    function getCmdKFeatureUsage(data) {
        return data.reduce((sum, row) => sum + (parseFloat(row[COLUMN_CMDK_USAGES]) || 0), 0);
    }

    /**
     * Calculates top N users by a rate (accepted / total).
     * @param {Array<Object>} aggregatedData - The aggregated user data.
     * @param {string} acceptedColKey - The key for the 'accepted' count column for each user.
     * @param {string} totalColKey - The key for the 'total' count column for each user.
     * @param {number} [topN=10] - The number of top users to return.
     * @returns {Array<Object>} Sorted list of top N users with email, rate, acceptedCount, totalCount.
     */
    function getTopUsersByRate(aggregatedData, acceptedColKey, totalColKey, topN = 10) {
        const usersWithRates = aggregatedData.map(user => {
            const acceptedCount = parseFloat(user[acceptedColKey]) || 0;
            const totalCount = parseFloat(user[totalColKey]) || 0;
            let rate = null;

            if (totalCount > 0) {
                rate = acceptedCount / totalCount;
            } else if (acceptedCount === 0) { // 0 accepted, 0 total means 0% rate for this context
                rate = 0;
            }
            // If rate is still null (e.g., accepted > 0, total = 0, indicates a data issue), it will be filtered.

            return {
                email: user[COLUMN_EMAIL],
                rate: rate,
                acceptedCount: acceptedCount,
                totalCount: totalCount,
                userId: user[COLUMN_USER_ID] // For potential tie-breaking or detailed views
            };
        }).filter(user => user.rate !== null); // Filter out users where rate couldn't be determined

        // Sort by rate (descending), then by acceptedCount (descending) as a tie-breaker
        usersWithRates.sort((a, b) => {
            // Handle null rates if any somehow pass the filter (should not happen with current logic)
            if (a.rate === null && b.rate === null) return 0;
            if (a.rate === null) return 1; // Sort nulls (N/A) to the bottom
            if (b.rate === null) return -1;

            if (b.rate !== a.rate) return b.rate - a.rate;
            return b.acceptedCount - a.acceptedCount; // Higher accepted count is better for ties
        });

        return usersWithRates.slice(0, topN);
    }

    /**
     * Calculates the popularity breakdown of different models among active users.
     * Popularity is defined as the percentage of active users who have used a particular model.
     * @param {Array<Object>} data - The aggregated data.
     * @returns {Array<{model: string, percentage: number}>} An array of objects, each with `model` name and its usage `percentage`. Sorted by percentage in descending order. Returns an empty array if no active users.
     */
    function getModelPopularityBreakdown(data) {
        const activeUsers = data.filter(row => row[COLUMN_IS_ACTIVE] === 'TRUE');
        if (activeUsers.length === 0) return [];

        const modelCounts = new Map();
        activeUsers.forEach(user => {
            const modelsStr = user[COLUMN_MOST_USED_MODEL];
            if (modelsStr) {
                const models = modelsStr.split(',').map(m => m.trim()).filter(m => m !== '');
                models.forEach(model => {
                    if (model) {
                        modelCounts.set(model, (modelCounts.get(model) || 0) + 1);
                    }
                });
            }
        });

        const popularity = [];
        modelCounts.forEach((count, model) => {
            popularity.push({ model, percentage: (count / activeUsers.length) * 100 });
        });

        return popularity.sort((a, b) => b.percentage - a.percentage);
    }

    /**
     * Identifies the top 10 power users based on their total AI lines added (suggested + accepted).
     * @param {Array<Object>} data - The aggregated data.
     * @returns {Array<{email: string, aiLines: number}>} An array of the top 10 users (or fewer if less than 10 users exist), sorted by `aiLines` in descending order. Each object contains `email` and `aiLines`.
     */
    function getTop10PowerUsers(data) {
        const usersWithAILines = data.map(row => ({
            email: row[COLUMN_EMAIL],
            aiLines: (parseFloat(row[COLUMN_CHAT_SUGGESTED_LINES_ADDED]) || 0) + (parseFloat(row[COLUMN_CHAT_ACCEPTED_LINES_ADDED]) || 0)
        }));

        usersWithAILines.sort((a, b) => b.aiLines - a.aiLines);
        return usersWithAILines.slice(0, 10);
    }

    /**
     * Calculates the rank of a given user based on their AI lines added, compared to all users.
     * @param {Object} currentUserData - The data object for the user whose rank is to be calculated.
     * @param {Array<Object>} allUsersData - An array of data objects for all users (typically originalAggregatedData).
     * @returns {string|number} The rank of the user (e.g., 1, 2, 24), or 'N/A' if not applicable.
     */
    function calculateUserPowerRank(currentUserData, allUsersData) {
        const getUserAILines = (user) => (parseFloat(user[COLUMN_CHAT_SUGGESTED_LINES_ADDED]) || 0) + (parseFloat(user[COLUMN_CHAT_ACCEPTED_LINES_ADDED]) || 0);

        const currentUserAILines = getUserAILines(currentUserData);

        const allUsersWithAILines = allUsersData.map(user => ({
            email: user[COLUMN_EMAIL], // For identification
            aiLines: getUserAILines(user)
        }));

        // Sort by AI lines descending, then by email ascending for stable sort
        allUsersWithAILines.sort((a, b) => {
            if (b.aiLines !== a.aiLines) {
                return b.aiLines - a.aiLines;
            }
            // For users with the same AI lines, maintain a consistent order, e.g., by email
            return a.email.localeCompare(b.email);
        });

        // Find the 0-based index, rank is 1-based
        const rank = allUsersWithAILines.findIndex(user => user.email === currentUserData[COLUMN_EMAIL]) + 1;
        
        return rank > 0 ? rank : 'N/A'; // Should always find the user if currentUserData is from allUsersData
    }

    /**
     * Renders a simple text metric within a chart-like container for visual consistency.
     * @param {string} title - The title of the metric.
     * @param {string} valueText - The text value to display (e.g., "Active", "GPT-4", "#15").
     * @param {string} chartId - A unique base ID for the elements (wrapper will get _wrapper suffix).
     * @param {HTMLElement} containerElement - The HTML grid element to append the metric display to.
     * @returns {HTMLElement | null} The created wrapper element, or null if creation fails.
     */
    function renderTextMetricInChartSlot(title, valueText, chartId, containerElement) {
        if (!containerElement) {
            console.error(`Error: Container element for text metric '${title}' not found.`);
            return null;
        }

        const chartWrapper = document.createElement('div');
        chartWrapper.className = 'chart-container'; // Use existing class for styling
        if (chartId) chartWrapper.id = chartId + '_wrapper';

        const titleEl = document.createElement('h4');
        // titleEl.style.textAlign = 'center'; // Already handled by .chart-container h4 CSS
        // titleEl.style.marginTop = '0';
        // titleEl.style.marginBottom = '15px';
        // titleEl.style.color = '#391085';
        titleEl.textContent = title;
        chartWrapper.appendChild(titleEl);

        const valueEl = document.createElement('p');
        valueEl.textContent = valueText;
        valueEl.style.textAlign = 'center';
        valueEl.style.fontSize = '2.2em'; // Larger font for the value
        valueEl.style.fontWeight = 'bold';
        valueEl.style.color = '#FF00A8'; // Lemonade Pink, consistent with metric values
        valueEl.style.marginTop = '30px'; // More space between title and value
        valueEl.style.lineHeight = '1.2';
        chartWrapper.appendChild(valueEl);

        containerElement.appendChild(chartWrapper);
        return chartWrapper;
    }

    /**
     * Main function to display all dashboard metrics and charts on the UI.
     * It clears previous content and populates the key metrics grid and the charts grid based on the provided aggregated data.
     * @param {Array<Object>} aggregatedData - The fully aggregated and formatted data ready for display.
     * @param {HTMLElement} chartsContainerEl - The HTML element that will serve as the container for the main charts grid (e.g., a `div` with class 'charts-grid').
     */
    function displayDashboardMetrics(aggregatedData, chartsContainerEl, rawUserData = null) {
        console.log('displayDashboardMetrics called.');
        // const dashboardArea = document.getElementById('dashboardArea'); // Already cached as dashboardAreaEl
        console.log('dashboardArea element (cached as dashboardAreaEl):', dashboardAreaEl);

        if (dashboardAreaEl) {
            dashboardAreaEl.style.display = 'block'; // Ensure it's visible
            console.log('dashboardArea display set to block');
        } else {
            console.error('CRITICAL: dashboardAreaEl not found! Dashboard cannot be displayed.');
            return;
        }
        if (!chartsContainerEl) { // Should be chartsDisplayEl
            console.error('CRITICAL: chartsContainerEl (expected chartsDisplayEl) not found or passed incorrectly!');
            return;
        }

        // Clear previous content from key metrics and charts grids
        if (keyMetricsDisplayEl) keyMetricsDisplayEl.innerHTML = '';
        chartsContainerEl.innerHTML = '';

        const isSingleUserView = aggregatedData.length === 1;
        const singleUserData = isSingleUserView ? aggregatedData[0] : null;

        // --- Key Metrics Grid Population ---
        if (keyMetricsDisplayEl) {
            const activeDevCount = getActiveDeveloperCount(aggregatedData);
            renderKeyValueMetric(
                'Active Developers',
                activeDevCount,
                '',
                'Count of developers marked as active.',
                keyMetricsDisplayEl
            );

            const totalAILinesAdded = getTotalAIGeneratedLinesAdded(aggregatedData);
            renderKeyValueMetric(
                'Total AI-Generated Lines',
                totalAILinesAdded,
                'lines',
                'Sum of all suggested and accepted AI lines.',
                keyMetricsDisplayEl
            );

            const avgLinesPerActiveUser = (activeDevCount > 0 && totalAILinesAdded !== null) ? (totalAILinesAdded / activeDevCount) : 0;
            renderKeyValueMetric(
                'Avg. Lines / Active User',
                parseFloat(avgLinesPerActiveUser.toFixed(1)),
                'lines',
                'Total AI lines divided by active developers.',
                keyMetricsDisplayEl
            );

            const promptActivityIndex = getPromptActivityIndex(aggregatedData);
            renderKeyValueMetric(
                'Prompt Interactions',
                promptActivityIndex,
                'interactions',
                'Total Ask, Edit, and Agent requests.',
                keyMetricsDisplayEl
            );

            const cmdKUsage = getCmdKFeatureUsage(aggregatedData);
            renderKeyValueMetric(
                'Cmd+K Usages',
                cmdKUsage,
                'usages',
                'Total uses of the Cmd+K feature.',
                keyMetricsDisplayEl
            );
        } else {
            console.warn("keyMetricsDisplayEl not found. Key metrics card display will be skipped.");
        }


        // --- Charts Grid Population & Inactive Users List ---
        const inactiveUsersDisplayEl = document.getElementById('inactiveUsersDisplay');

        if (isSingleUserView && singleUserData) {
            // --- Single User View Logic for Charts Area ---

            // 1. User Status
            const userStatus = singleUserData[COLUMN_IS_ACTIVE] === 'TRUE' ? 'Active' : 'Inactive';
            renderTextMetricInChartSlot('User Status', userStatus, 'userStatusMetric', chartsContainerEl);

            // 2. Model Popularity (Most Used Model for the user)
            const mostUsedModel = singleUserData[COLUMN_MOST_USED_MODEL] || 'N/A';
            renderTextMetricInChartSlot('Most Used Model', mostUsedModel, 'userModelMetric', chartsContainerEl);
            
            // 3. User Power Rank
            // Ensure originalAggregatedData is available (it's a global variable initialized earlier)
            if (originalAggregatedData && originalAggregatedData.length > 0) {
                const powerRank = calculateUserPowerRank(singleUserData, originalAggregatedData);
                renderTextMetricInChartSlot('Power User Rank', `#${powerRank}`, 'userPowerRankMetric', chartsContainerEl);
            } else {
                renderTextMetricInChartSlot('Power User Rank', 'N/A', 'userPowerRankMetric', chartsContainerEl);
                console.warn("Original aggregated data not available for power rank calculation.");
            }

            // 4. AI Suggestion Rate for the user
            const userTotalAccepts = parseFloat(singleUserData[COLUMN_CHAT_TOTAL_ACCEPTS]) || 0;
            const userTotalApplies = parseFloat(singleUserData[COLUMN_CHAT_TOTAL_APPLIES]) || 0;
            const userAcceptanceRate = userTotalApplies > 0 ? (userTotalAccepts / userTotalApplies) : null;
            
            if (userAcceptanceRate !== null) {
                renderPercentageGaugeChart(
                    'Your AI Suggestion Rate', 
                    userAcceptanceRate, 
                    'Your (Chat Total Accepts) / (Chat Total Applies)', 
                    'userAcceptanceRateChart', 
                    chartsContainerEl
                );
            } else {
                renderTextMetricInChartSlot('AI Suggestion Rate', 'N/A', 'userAcceptanceRateChart', chartsContainerEl);
            }

            // 5. "Tab" Suggestion Rate for the user
            const userTabsAccepted = parseFloat(singleUserData[COLUMN_TABS_ACCEPTED]) || 0;
            const userChatTabsShown = parseFloat(singleUserData[COLUMN_CHAT_TABS_SHOWN]) || 0;
            const userTabConversionRate = userChatTabsShown > 0 ? (userTabsAccepted / userChatTabsShown) : null;

            if (userTabConversionRate !== null) {
                renderPercentageGaugeChart(
                    'Your "Tab" Suggestion Rate', 
                    userTabConversionRate, 
                    'Your (Tabs Accepted) / (Chat Tabs Shown)', 
                    'userTabConversionChart', 
                    chartsContainerEl
                );
            } else {
                renderTextMetricInChartSlot('"Tab" Suggestion Rate', 'N/A', 'userTabConversionChart', chartsContainerEl);
            }

            // Add Daily Interactions Chart for single user
            if (rawUserData) {
                const dailyInteractions = getDailyPromptInteractions(rawUserData);
                renderDailyInteractionsChart(dailyInteractions, 'userDailyInteractionsChart', chartsContainerEl);
            } else {
                const dailyInteractions = getDailyPromptInteractions(currentFilteredData);
                renderDailyInteractionsChart(dailyInteractions, 'userDailyInteractionsChart', chartsContainerEl);
            }
            
            // 6. Hide Inactive Users Section
            if (inactiveUsersDisplayEl) {
                inactiveUsersDisplayEl.style.display = 'none';
            }

        } else {
            // --- Multi-User (Default) View Logic for Charts Area ---
            const activeDevCountForChart = getActiveDeveloperCount(aggregatedData);
            const inactiveDevCount = getInactiveDeveloperCount(aggregatedData);
            renderDevStatusChart(activeDevCountForChart, inactiveDevCount, chartsContainerEl);

            const acceptanceRate = getAcceptanceRateOfAISuggestions(aggregatedData);
            if (acceptanceRate !== null) {
                renderPercentageGaugeChart(
                    'AI Suggestion Acceptance Rate',
                    acceptanceRate,
                    'Σ(Chat Total Accepts) ÷ Σ(Chat Total Applies)',
                    'acceptanceRateChart',
                    chartsContainerEl
                );
            } else {
                // Fallback for charts grid if data is N/A - using renderTextMetricInChartSlot for consistency
                renderTextMetricInChartSlot('AI Suggestion Acceptance Rate', 'N/A', 'acceptanceRateChart_fallback', chartsContainerEl);
            }

            const tabConversionRate = getTabConversionRate(aggregatedData);
            if (tabConversionRate !== null) {
                renderPercentageGaugeChart(
                    '"Tab" Suggestion Conversion Rate',
                    tabConversionRate,
                    'Σ(Tabs Accepted) ÷ Σ(Chat Tabs Shown)',
                    'tabConversionChart',
                    chartsContainerEl
                );
            } else {
                renderTextMetricInChartSlot('"Tab" Suggestion Conversion Rate', 'N/A', 'tabConversionChart_fallback', chartsContainerEl);
            }

            const modelPopularity = getModelPopularityBreakdown(aggregatedData);
            renderModelPopularityChart(modelPopularity, chartsContainerEl);

            const topPowerUsers = getTop10PowerUsers(aggregatedData);
            renderTopPowerUsersChart(topPowerUsers, chartsContainerEl);

            const topAiSuggestionUsers = getTopUsersByRate(aggregatedData, COLUMN_CHAT_TOTAL_ACCEPTS, COLUMN_CHAT_TOTAL_APPLIES, 10);
            renderTopUsersRateChart(
                'Top 10 AI Suggestion Users',
                'AI Suggestion Rate',
                topAiSuggestionUsers,
                'topAiSuggestionUsersChart',
                chartsContainerEl,
                '#4CAF50', 
                '#388E3C'
            );

            const topTabSuggestionUsers = getTopUsersByRate(aggregatedData, COLUMN_TABS_ACCEPTED, COLUMN_CHAT_TABS_SHOWN, 10);
            renderTopUsersRateChart(
                'Top 10 "Tab" Suggestion Users',
                '"Tab" Suggestion Rate',
                topTabSuggestionUsers,
                'topTabSuggestionUsersChart',
                chartsContainerEl,
                '#2196F3', 
                '#1976D2'
            );

            // Add Daily Interactions Chart for all users
            const dailyInteractions = getDailyPromptInteractions(rawCsvData);
            renderDailyInteractionsChart(dailyInteractions, 'allUsersDailyInteractionsChart', chartsContainerEl);

            // Show and populate Inactive Users Section
            if (inactiveUsersDisplayEl) {
                inactiveUsersDisplayEl.style.display = 'block'; // Or its CSS default
                const inactiveUsers = getInactiveUsersList(aggregatedData);
                renderInactiveUsersList(inactiveUsers);
            }
        }

        console.log('Dashboard metrics and charts display updated.');
    }

    /**
     * Renders a bar chart for top N users based on a calculated rate (e.g., acceptance rate, conversion rate).
     * Each bar represents a user (email prefix), and its length corresponds to their rate.
     * @param {string} chartTitle - The title for the chart.
     * @param {string} dataLabelPrefix - Prefix for the tooltip label (e.g., "Acceptance Rate").
     * @param {Array<Object>} usersData - Array of user objects, must contain `email`, `rate`, `acceptedCount`, `totalCount`.
     * @param {string} chartId - The unique ID for the canvas element.
     * @param {HTMLElement} containerElement - The HTML element to append the chart's container to.
     * @param {string} barColor - The color for the bars in the chart (hex or rgba).
     * @param {string} borderColor - The border color for the bars.
     */
    function renderTopUsersRateChart(chartTitle, dataLabelPrefix, usersData, chartId, containerElement, barColor = '#FFA726', borderColor = '#FB8C00') {
        const ctx = createChartCanvas(chartTitle, chartId, containerElement);
        if (!ctx) return;

        if (!usersData || usersData.length === 0) {
            const chartWrapper = ctx.canvas.parentElement;
            if (chartWrapper) {
                chartWrapper.innerHTML = ''; // Clear existing content
                const titleEl = document.createElement('h4');
                titleEl.textContent = chartTitle; // chartTitle is from hardcoded strings in current use
                const pEl = document.createElement('p');
                pEl.style.textAlign = 'center';
                pEl.style.paddingTop = '20px';
                pEl.textContent = 'No data available for this chart.';
                chartWrapper.appendChild(titleEl);
                chartWrapper.appendChild(pEl);
            }
            return;
        }

        const labels = usersData.map(user => user.email.substring(0, user.email.indexOf('@') > 0 ? user.email.indexOf('@') : 20)); // Shorten email
        const data = usersData.map(user => user.rate !== null ? user.rate * 100 : 0); // Convert rate to percentage for display

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: dataLabelPrefix,
                    data: data,
                    backgroundColor: barColor,
                    borderColor: borderColor,
                    borderWidth: 1
                }]
            },
            options: {
                indexAxis: 'y',
                scales: {
                    x: {
                        beginAtZero: true,
                        max: 100, // Rates are 0-100%
                        ticks: {
                            callback: function(value) {
                                return value + '%'; // Add percentage sign to x-axis ticks
                            }
                        }
                    }
                },
                plugins: {
                    title: {
                        display: false, // Handled by createChartCanvas
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const user = usersData[context.dataIndex];
                                const rateFormatted = formatPercentage(user.rate);
                                return `${dataLabelPrefix}: ${rateFormatted} (${user.acceptedCount}/${user.totalCount})`;
                            }
                        }
                    }
                }
            }
        });
    }

    // --- END OF UTILITY FUNCTIONS ---

    // --- NEW Search and Autocomplete Functions ---

    /**
     * Handles input events on the user search input field.
     * Triggers the display of autocomplete suggestions.
     * @param {Event} event - The input event.
     */
    function handleUserSearchInput(event) {
        const searchTerm = event.target.value.trim();
        if (searchTerm.length > 0) {
            displayAutocompleteSuggestions(searchTerm);
        } else {
            if (autocompleteSuggestionsEl) autocompleteSuggestionsEl.style.display = 'none';
            // If search is cleared, show all users again
            if (JSON.stringify(currentFilteredData) !== JSON.stringify(originalAggregatedData)) {
                clearUserFilterAndDisplayAll();
            }
        }
    }

    /**
     * Filters users based on the search term and displays autocomplete suggestions.
     * @param {string} searchTerm - The text entered by the user.
     */
    function displayAutocompleteSuggestions(searchTerm) {
        if (!autocompleteSuggestionsEl || !originalAggregatedData) return;

        const filteredUsers = originalAggregatedData.filter(user => 
            user[COLUMN_EMAIL].toLowerCase().includes(searchTerm.toLowerCase())
        );

        autocompleteSuggestionsEl.innerHTML = ''; // Clear previous suggestions

        if (filteredUsers.length > 0) {
            filteredUsers.slice(0, 10).forEach(user => { // Limit to 10 suggestions
                const suggestionDiv = document.createElement('div');
                suggestionDiv.textContent = user[COLUMN_EMAIL];
                suggestionDiv.style.padding = '8px';
                suggestionDiv.style.cursor = 'pointer';
                suggestionDiv.addEventListener('mouseenter', () => suggestionDiv.style.backgroundColor = '#f0f0f0');
                suggestionDiv.addEventListener('mouseleave', () => suggestionDiv.style.backgroundColor = 'white');
                // suggestionDiv.addEventListener('click', () => selectUserFromAutocomplete(user[COLUMN_EMAIL]));
                autocompleteSuggestionsEl.appendChild(suggestionDiv);
            });
            autocompleteSuggestionsEl.style.display = 'block';
        } else {
            autocompleteSuggestionsEl.style.display = 'none';
        }
    }

    /**
     * Handles the click event on an autocomplete suggestion.
     * Filters data for the selected user and updates the dashboard.
     * @param {Event} event - The click event from the suggestions container.
     */
    function handleSuggestionClick(event) {
        if (event.target && event.target.nodeName === 'DIV' && event.target.textContent) {
            const selectedEmail = event.target.textContent;
            if (userSearchInputEl) {
                userSearchInputEl.value = selectedEmail; // Populate search bar with selected email
            }
            if (autocompleteSuggestionsEl) {
                autocompleteSuggestionsEl.style.display = 'none'; // Hide suggestions
            }
            filterAndDisplayUser(selectedEmail);
        }
    }

    /**
     * Filters the aggregated data for a specific user and re-renders the dashboard.
     * @param {string} email - The email of the user to filter by.
     */
    function filterAndDisplayUser(email) {
        if (!originalAggregatedData) return;
        
        currentFilteredData = originalAggregatedData.filter(user => user[COLUMN_EMAIL] === email);
        
        if (currentFilteredData.length > 0) {
            // Get raw data for daily interactions chart
            const userRawData = rawCsvData.filter(row => row[COLUMN_EMAIL] === email);
            
            if (chartsDisplayEl) {
                displayDashboardMetrics(currentFilteredData, chartsDisplayEl, userRawData);
            }
            if (reportDatesDivEl) {
                reportDatesDivEl.innerHTML = '';
                const textPart = document.createTextNode('Displaying data for: ');
                const emailStrongPart = document.createElement('strong');
                emailStrongPart.textContent = email;
                reportDatesDivEl.appendChild(textPart);
                reportDatesDivEl.appendChild(emailStrongPart);
                reportDatesDivEl.style.display = 'block';
            }
        } else {
            console.warn(`[DEBUG] No data found for user: ${email} after filtering.`);
            resetUI(`No data found for selected user: ${email}.`);
        }
    }

    /**
     * Clears any active user filter and displays data for all users.
     */
    function clearUserFilterAndDisplayAll() {
        console.log('[DEBUG] Clearing user filter, displaying all data.');
        currentFilteredData = [...originalAggregatedData];
        if (userSearchInputEl) userSearchInputEl.value = ''; // Clear search input
        if (autocompleteSuggestionsEl) autocompleteSuggestionsEl.style.display = 'none';

        if (chartsDisplayEl) {
            displayDashboardMetrics(currentFilteredData, chartsDisplayEl);
        }
        // Restore original report dates display logic
        const { reportStartDate, reportEndDate } = aggregateCsvData(originalAggregatedData); // This is inefficient, should store these
        // For now, let's just revert to a generic message or re-calculate if we must.
        // Better: Store original reportStartDate and reportEndDate in processParsedData and reuse here.
        // Simplified for now:
        if(reportDatesDivEl && originalAggregatedData.length > 0) {
            // This part needs access to overall report dates, which were calculated in processParsedData
            // Simplification: re-calculate (not ideal) or retrieve stored values if available
            // Assuming `storedReportStartDate` and `storedReportEndDate` were saved globally from `processParsedData`
            // For this example, I'll call a simplified date display or let displayReportDates handle it if it can
            // The original call in processParsedData used: displayReportDates(reportStartDate, reportEndDate);
            // We need to get those original dates. The easiest is to re-calculate from originalAggregatedData if not stored.

            // Minimal re-calc for dates (Not the most efficient, but works for now):
            let overallOldestDate = null;
            let overallNewestDate = null;
            if (originalAggregatedData.length > 0) {
                // Ensure that 'Dates' property exists and is a string before splitting
                const allDates = originalAggregatedData.flatMap(entry => {
                    if (entry && typeof entry.Dates === 'string') {
                        const parts = entry.Dates.split(' / ');
                        return parts.map(d => parseDate(d)); // parseDate needs to handle DD-MM-YYYY
                    }
                    return [];
                }).filter(d => d instanceof Date && !isNaN(d)); // Ensure only valid dates proceed

                if (allDates.length > 0) {
                    overallOldestDate = new Date(Math.min(...allDates.map(d => d.getTime())));
                    overallNewestDate = new Date(Math.max(...allDates.map(d => d.getTime())));
                }
            }
            const formattedOverallOldestDate = overallOldestDate ? formatDateWithoutTime(overallOldestDate) : 'NODATE';
            const formattedOverallNewestDate = overallNewestDate ? formatDateWithoutTime(overallNewestDate) : 'NODATE';
            displayReportDates(formattedOverallOldestDate, formattedOverallNewestDate);
        } else if (reportDatesDivEl) {
            reportDatesDivEl.style.display = 'none';
        }
    }

    // --- END OF NEW Search and Autocomplete Functions ---

    /**
     * Calculates daily prompt interactions (Ask + Edit + Agent requests) for a user.
     * @param {Array<Object>} userData - The data rows for a specific user.
     * @returns {Array<{date: Date, interactions: number}>} Array of daily interactions, sorted by date.
     */
    function getDailyPromptInteractions(userData) {
        const dailyData = new Map();

        userData.forEach(row => {
            const dateStr = row[COLUMN_DATE];
            const date = parseDate(dateStr);
            if (!date) return;

            const ask = parseFloat(row[COLUMN_ASK_REQUESTS]) || 0;
            const edit = parseFloat(row[COLUMN_EDIT_REQUESTS]) || 0;
            const agent = parseFloat(row[COLUMN_AGENT_REQUESTS]) || 0;
            const totalInteractions = ask + edit + agent;

            const formattedDate = formatDateWithoutTime(date);
            dailyData.set(formattedDate, (dailyData.get(formattedDate) || 0) + totalInteractions);
        });

        // Convert to array and sort by date
        return Array.from(dailyData.entries())
            .map(([dateStr, interactions]) => ({
                date: parseDate(dateStr),
                interactions: interactions
            }))
            .sort((a, b) => a.date - b.date);
    }

    /**
     * Renders a line chart showing daily prompt interactions.
     * @param {Array<{date: Date, interactions: number}>} dailyData - Array of daily interaction counts.
     * @param {string} chartId - The ID for the chart canvas.
     * @param {HTMLElement} containerElement - The container to render the chart in.
     */
    function renderDailyInteractionsChart(dailyData, chartId, containerElement) {
        const ctx = createChartCanvas('Daily Prompt Interactions', chartId, containerElement);
        if (!ctx) return;

        if (!dailyData || dailyData.length === 0) {
            const chartWrapper = ctx.canvas.parentElement;
            if (chartWrapper) {
                chartWrapper.innerHTML = '<h4>Daily Prompt Interactions</h4><p style="text-align: center; padding-top: 20px;">No daily interaction data available.</p>';
            }
            return;
        }

        new Chart(ctx, {
            type: 'line',
            data: {
                labels: dailyData.map(d => formatDateWithoutTime(d.date)),
                datasets: [{
                    label: 'Prompt Interactions',
                    data: dailyData.map(d => d.interactions),
                    borderColor: '#FF00A8',
                    backgroundColor: 'rgba(255, 0, 168, 0.1)',
                    borderWidth: 2,
                    fill: true,
                    tension: 0.1
                }]
            },
            options: {
                responsive: true,
                scales: {
                    x: {
                        display: true,
                        title: {
                            display: true,
                            text: 'Date'
                        }
                    },
                    y: {
                        display: true,
                        title: {
                            display: true,
                            text: 'Interactions'
                        },
                        beginAtZero: true
                    }
                },
                plugins: {
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                return `Interactions: ${context.parsed.y}`;
                            }
                        }
                    }
                }
            }
        });
    }

})(); // End of IIFE