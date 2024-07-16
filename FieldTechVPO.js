document.addEventListener("DOMContentLoaded", async function () {
    console.log('DOM fully loaded and parsed');

    const airtableApiKey = 'pata9Iv7DANqtJrgO.b308b33cd0f323601f3fb580aac0d333ca1629dd26c5ebe2e2b9f18143ccaa8e';
    const airtableBaseId = 'appQDdkj6ydqUaUkE';
    const airtableTableName = 'tblO72Aw6qplOEAhR';
    const airtableEndpoint = `https://api.airtable.com/v0/${airtableBaseId}/${airtableTableName}`;

    const dropboxAccessToken = 'sl.B5IoIcnEbvoYHLsRM5L721cg-owwK2WnS9GI1Cw6ZA1glc6JeClGtpte2tlYus4mK9lavdMFJMT-aTzIG3WdtolD3Xjlfe2nTOGK6Ylbl6vuIqqr4aOI6DDyn9NXt0uS1oPSwG3XWgQjkOM1wr5FoAE';
    axios.defaults.headers.common['Authorization'] = `Bearer ${airtableApiKey}`;

    
    let allRecords = [];

    async function fetchAllRecords() {
        console.log('Fetching all records from Airtable...');
        let records = [];
        let offset = null;

        do {
            const response = await fetch(`${airtableEndpoint}?${new URLSearchParams({ offset })}`, {
                headers: {
                    Authorization: `Bearer ${airtableApiKey}`
                }
            });

            const data = await response.json();
            records = records.concat(data.records.map(record => ({
                id: record.id,
                fields: record.fields,
                descriptionOfWork: record.fields['Description of Work']
            })));
            offset = data.offset;
        } while (offset);

        console.log(`Total records fetched: ${records.length}`);
        return records;
    }

    async function fetchUncheckedRecords() {
        try {
            showLoadingMessage();
            console.log('Fetching unchecked records from Airtable...');
            const filterByFormula = 'NOT({Field Tech Confirmed Job Complete})';
            let records = [];
            let offset = '';

            do {
                const response = await axios.get(`${airtableEndpoint}?filterByFormula=${encodeURIComponent(filterByFormula)}&offset=${offset}`);
                records = records.concat(response.data.records.map(record => ({
                    id: record.id,
                    fields: record.fields,
                    descriptionOfWork: record.fields['Description of Work']
                })));
                offset = response.data.offset || '';
            } while (offset);

            console.log(`Unchecked records fetched successfully: ${records.length} records`);
            allRecords = records;
            displayRecords(records);
        } catch (error) {
            console.error('Error fetching unchecked records:', error);
        } finally {
            hideLoadingMessage();
        }
    }

    function showLoadingMessage() {
        console.log('Showing loading message');
        document.getElementById('loadingMessage').innerText = 'Open VPOs are being loaded...';
        document.getElementById('loadingMessage').style.display = 'block';
        document.getElementById('searchButton').classList.add('hidden');
        document.getElementById('submitUpdates').classList.add('hidden');
        document.getElementById('searchBar').classList.add('hidden');
        document.getElementById('searchBarTitle').classList.add('hidden');
    }

    function hideLoadingMessage() {
        console.log('Hiding loading message');
        document.getElementById('loadingMessage').style.display = 'none';
        document.getElementById('searchButton').classList.remove('hidden');
        document.getElementById('submitUpdates').classList.remove('hidden');
        document.getElementById('searchBar').classList.remove('hidden');
        document.getElementById('searchBarTitle').classList.remove('hidden');
    }

    function displayRecords(records) {
        console.log('Displaying records...');
        const recordsContainer = document.getElementById('records');
        recordsContainer.innerHTML = '';

        if (records.length === 0) {
            recordsContainer.innerText = 'No records found.';
            console.log('No records found.');
            return;
        }

        records = sortRecordsWithSpecialCondition(records);

        const tableHeader = `
            <thead>
                <tr>
                    <th>Vanir Office</th>
                    <th>Job Name</th>
                    <th>Description of Work</th>
                    <th>Field Technician</th>
                    <th>Confirmed Complete</th>
                    <th>Completed Photo(s)</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        `;
        recordsContainer.innerHTML = tableHeader;
        const tableBody = recordsContainer.querySelector('tbody');

        records.forEach(record => {
            const recordRow = createRecordRow(record);
            tableBody.appendChild(recordRow);
        });

        console.log(`Total number of entries displayed: ${records.length}`);
    }

    function sortRecordsWithSpecialCondition(records) {
        console.log('Sorting records with special condition');
        return records.sort((a, b) => {
            const officeA = a.fields['static Vanir Office'] || '';
            const officeB = b.fields['static Vanir Office'] || '';
            const techA = a.fields['static Field Technician'] || '';
            const techB = b.fields['static Field Technician'] || '';

            if (officeA === 'Greensboro' && officeB === 'Greenville, SC') return -1;
            if (officeA === 'Greenville, SC' && officeB === 'Greensboro') return 1;

            const primarySort = officeA.localeCompare(officeB);
            if (primarySort !== 0) return primarySort;

            return techA.localeCompare(techB);
        });
    }

    function createRecordRow(record) {
        const recordRow = document.createElement('tr');
        const vanirOffice = record.fields['static Vanir Office'] || '';
        const jobName = record.fields['Job Name'] || '';
        const fieldTechnician = record.fields['static Field Technician'] || '';
        const fieldTechConfirmedComplete = record.fields['Field Tech Confirmed Job Complete'];
        const checkboxValue = fieldTechConfirmedComplete ? 'checked' : '';
        const descriptionOfWork = record.descriptionOfWork || '';

        recordRow.innerHTML = `
            <td>${vanirOffice}</td>
            <td>${jobName}</td>
            <td>${descriptionOfWork}</td>
            <td>${fieldTechnician}</td>
            <td>
                <label class="custom-checkbox">
                    <input type="checkbox" ${checkboxValue} data-record-id="${record.id}" data-initial-checked="${checkboxValue}">
                    <span class="checkmark"></span>
                </label>
            </td>
            <td>
                <input type="file" class="file-upload hidden" data-record-id="${record.id}">
            </td>
        `;

        const checkbox = recordRow.querySelector('input[type="checkbox"]');
        const fileInput = recordRow.querySelector('.file-upload');
        checkbox.addEventListener('change', handleCheckboxChange);
        fileInput.addEventListener('change', handleFileSelection);

        console.log(`Created row for record ID ${record.id}:`, record);
        return recordRow;
    }

    function handleCheckboxChange(event) {
        const checkbox = event.target;
        const recordId = checkbox.getAttribute('data-record-id');
        const isChecked = checkbox.checked;
        const fileInput = document.querySelector(`input.file-upload[data-record-id="${recordId}"]`);

        let updates = JSON.parse(localStorage.getItem('updates')) || {};

        if (isChecked) {
            updates[recordId] = true;
            fileInput.classList.remove('hidden');
            console.log(`Checkbox checked for record ID ${recordId}. Showing file input.`);
        } else {
            delete updates[recordId];
            fileInput.classList.add('hidden');
            console.log(`Checkbox unchecked for record ID ${recordId}. Hiding file input.`);
        }

        localStorage.setItem('updates', JSON.stringify(updates));
        console.log('Current updates:', updates);
    }

    function handleFileSelection(event) {
        const fileInput = event.target;
        const recordId = fileInput.getAttribute('data-record-id');
        const file = fileInput.files[0];

        if (!file) return;

        console.log(`File selected for record ID ${recordId}:`, file.name);

        let fileData = JSON.parse(localStorage.getItem('fileData')) || {};
        fileData[recordId] = file;
        localStorage.setItem('fileData', JSON.stringify(fileData));
    }

    async function submitUpdates() {
        console.log('Submitting updates...');
        let updates = JSON.parse(localStorage.getItem('updates')) || {};
        let fileData = JSON.parse(localStorage.getItem('fileData')) || {};
        let updateArray = Object.keys(updates).map(id => ({
            id: id,
            fields: {
                'Field Tech Confirmed Job Complete': updates[id],
                'Field Tech Confirmed Job Completed Date': new Date().toISOString(),
                'Completed Photo(s)': '',
            },
        }));

        if (updateArray.length === 0) {
            console.log('No changes to submit.');
            alert('No changes to submit.');
            return;
        }

        const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

        async function patchWithRetry(url, data, retries = 5) {
            let attempt = 0;
            let success = false;
            let response = null;

            while (attempt < retries && !success) {
                try {
                    response = await axios.patch(url, data);
                    success = true;
                } catch (error) {
                    if (error.response && error.response.status === 429) {
                        attempt++;
                        const waitTime = Math.pow(2, attempt) * 1000;
                        console.log(`Rate limit exceeded. Retrying in ${waitTime / 1000} seconds...`);
                        await delay(waitTime);
                    } else {
                        throw error;
                    }
                }
            }

            if (!success) {
                throw new Error('Max retries reached. Failed to patch data.');
            }

            return response;
        }

        try {
            const updatePromises = updateArray.map(update => {
                const recordId = update.id;
                const file = fileData[recordId];
                if (file) {
                    return uploadFileToDropbox(file)
                        .then(dropboxLink => {
                            update.fields['Completed Photo(s)'] = dropboxLink;
                            console.log(`Uploaded file for record ID ${recordId} to Dropbox. Link: ${dropboxLink}`);
                        })
                        .catch(error => {
                            console.error(`Error uploading file for record ID ${recordId} to Dropbox:`, error);
                        })
                        .finally(() => {
                            delete fileData[recordId];
                            localStorage.setItem('fileData', JSON.stringify(fileData));
                        });
                } else {
                    console.warn(`No file selected for record ID ${recordId}.`);
                    return Promise.resolve();
                }
            });

            showLoadingMessage();
            console.log('Submitting updates to Airtable...', updatePromises);
            await Promise.all(updatePromises.map(p => p.catch(e => e))); // Ensure all promises are resolved, even if some fail
            console.log('Records updated successfully');
            alert('Records updated successfully.');

            localStorage.removeItem('updates');
            document.getElementById('loadingMessage').innerText = 'Values have been submitted. Repopulating table...';
            await fetchUncheckedRecords();
        } catch (error) {
            console.error('Error updating records:', error);
            alert('Error updating records. Check the console for more details.');
        } finally {
            hideLoadingMessage();
        }
    }

    async function uploadFileToDropbox(file) {
        try {
            console.log(`Uploading file to Dropbox: ${file.name}`);
    
            // Obtain Dropbox access token
            const token = await getDropboxAccessToken();
    
            // Construct Dropbox API upload URL
            const url = 'https://content.dropboxapi.com/2/files/upload';
            const headers = {
                Authorization: `Bearer ${token}`,
                'Dropbox-API-Arg': JSON.stringify({
                    path: `/uploads/${file.name}`,
                    mode: 'add',
                    autorename: true,
                    mute: false
                }),
                'Content-Type': 'application/octet-stream',
            };
    
            // Perform the file upload to Dropbox
            const response = await fetch(url, {
                method: 'POST',
                headers: headers,
                body: file
            });
    
            if (!response.ok) {
                throw new Error(`Failed to upload file to Dropbox. Status: ${response.status} ${response.statusText}`);
            }
    
            const responseData = await response.json();
            console.log('Dropbox upload response:', responseData);
    
            return responseData.path_display;
        } catch (error) {
            console.error(`Error uploading file to Dropbox:`, error);
            throw error; // Rethrow the error to propagate it up the call stack
        }
    }

    function filterRecords() {
        const searchTerm = document.getElementById('searchBar').value.toLowerCase();
        const filteredRecords = allRecords.filter(record => {
            const vanirOffice = (record.fields['static Vanir Office'] || '').toLowerCase();
            const jobName = (record.fields['Job Name'] || '').toLowerCase();
            const fieldTechnician = (record.fields['static Field Technician'] || '').toLowerCase();
            return vanirOffice.includes(searchTerm) || jobName.includes(searchTerm) || fieldTechnician.includes(searchTerm);
        });
        displayRecords(filteredRecords);
    }

    fetchAllRecords()
        .then(records => {
            console.log('Total records fetched:', records.length);
        })
        .catch(error => {
            console.error('Error fetching records:', error);
        });

    fetchUncheckedRecords();

    document.getElementById('submitUpdates').addEventListener('click', submitUpdates);
    document.getElementById('searchButton').addEventListener('click', filterRecords);
});