module.exports = async function (tp, jiraQuery, accountAliases) {
  // Function to update the Kanban file with Jira tasks
  async function updateKanbanWithJiraTasks() {
    // Get the currently opened file
    const activeFile = app.workspace.getActiveFile();
    if (!activeFile) {
      console.error('No file is currently open.');
      return;
    }

    // Read the content of the active file
    let kanbanContent = await app.vault.read(activeFile);

    // Check if the file has the frontmatter variable `autoUpdateKanban: true`
    const frontmatter = app.metadataCache.getFileCache(activeFile)?.frontmatter;
    const isAutoUpdateEnabled = frontmatter?.autoUpdateKanban === true;

    // If the frontmatter variable is not set, add it
    if (!isAutoUpdateEnabled) {
      console.log('First run detected. Adding frontmatter variable.');

      // Check if frontmatter already exists
      const hasExistingFrontmatter = kanbanContent.startsWith('---\n');

      if (hasExistingFrontmatter) {
        // Append `autoUpdateKanban: true` to existing frontmatter
        const frontmatterEndIndex = kanbanContent.indexOf('---\n', 4); // Find the end of the frontmatter block
        if (frontmatterEndIndex !== -1) {
          const existingFrontmatter = kanbanContent.slice(0, frontmatterEndIndex);
          const restOfContent = kanbanContent.slice(frontmatterEndIndex);
          kanbanContent = `${existingFrontmatter}autoUpdateKanban: true\n${restOfContent}`;
        }
      } else {
        kanbanContent = `---\nkanban-plugin: board\nautoUpdateKanban: true\n---\n${kanbanContent}`;
      }

      // Write the updated content back to the active file
      await app.vault.modify(activeFile, kanbanContent);
      return;
    }

    // Fetch Jira tasks based on the query and account aliases
    let jiraTasks = [];
    for (const accountAlias of accountAliases) {
      const tasks = await fetchJiraTasks(jiraQuery, accountAlias);
      jiraTasks = jiraTasks.concat(tasks);
    }

    // Dynamically pull headers from the file content
    const headers = extractHeadersFromContent(kanbanContent);

    // Process each Jira task
    for (const task of jiraTasks) {
      const status = task.fields.status.name;
      const taskDescription = `- [ ] [[${task.key}]]\n`;
      const fileName = `${task.key}.md`;
      const fileExists = tp.file.find_tfile(fileName);

      let fileDisplayName;
      if (fileExists) {
        fileDisplayName = fileExists.basename;
      } else {
        fileDisplayName = await tp.file.create_new(tp.file.find_tfile("Jira Ticket"), task.key).basename;
      }

      // Check if the task already exists in the file
      if (kanbanContent.includes(`${task.key}`)) {
        continue;
      }

      // Find the matching header
      const headerIndex = headers.indexOf(status);
      console.log(`headerIndex of ${status}: ${headerIndex}`);
      const targetHeader = headerIndex !== -1 ? headers[headerIndex] : 'To Sort';

      // Find the position of the target header
      const headerRegex = new RegExp(`##\\s+${targetHeader}\\n`, 'g');
      const headerMatch = kanbanContent.match(headerRegex);

      if (headerMatch) {
        // Find the position of the header in the content
        const headerPosition = kanbanContent.indexOf(`## ${targetHeader}\n`);

        // Check if there is a `**Complete**` line below the header
        const completeMarker = `**Complete**`;
        const completeMarkerPosition = kanbanContent.indexOf(completeMarker, headerPosition);
        let insertPosition = 0;
        if (completeMarkerPosition !== -1) {
          // If `**Complete**` exists below the header, insert the task below it
          insertPosition = completeMarkerPosition + completeMarker.length + 1; // +1 for the newline
        } else {
          // If `**Complete**` does not exist, insert the task directly below the header
          insertPosition = headerPosition + `## ${targetHeader}\n`.length;
        }

        // Insert the task at the correct position
        kanbanContent = kanbanContent.slice(0, insertPosition) + taskDescription + kanbanContent.slice(insertPosition);
      } else {
        // If the header does not exist, add the header and the task
        kanbanContent += `\n## ${targetHeader}\n${taskDescription}`;
      }
    }

    // Write the updated content back to the active file
    await app.vault.modify(activeFile, kanbanContent);
  }

  // Function to fetch Jira tasks
  async function fetchJiraTasks(query, accountAlias) {
    // Use the Obsidian Jira Issue plugin to search for tasks
    const jiraIssues = await searchJiraIssues(query, accountAlias);
    return jiraIssues;
  }

  // Function to search Jira issues
  async function searchJiraIssues(query, accountAlias) {
    // Ensure the Jira Issue plugin is loaded
    if (typeof $ji === 'undefined') {
      console.error('Obsidian Jira Issue plugin is not loaded.');
      return [];
    }

    const account = $ji.account.getAccountByAlias(accountAlias);

    // Call the plugin's getSearchResults method (correct method according to documentation)
    const searchResults = await $ji.defaulted.getSearchResults(query, { account });
    console.log(searchResults);
    // Format the results into a consistent structure
    return searchResults.issues.map(issue => ({
      key: issue.key,
      fields: {
        summary: issue.fields.summary,
        status: { name: issue.fields.status.name }
      }
    }));
  }

  // Function to extract headers from the file content
  function extractHeadersFromContent(content) {
    const headerRegex = /##\s+(.+)\n/g;
    const headers = [];
    let match;
    while ((match = headerRegex.exec(content)) !== null) {
      headers.push(match[1]);
    }
    return headers;
  }

  // Function to check if a file exists
  async function fileExists(fileName) {
    const file = app.vault.getAbstractFileByPath(fileName);
    return file !== null;
  }

  // Expose functions to Templater
  return {
    updateKanbanWithJiraTasks,
    fetchJiraTasks,
    searchJiraIssues,
    extractHeadersFromContent,
    fileExists
  };
};