module.exports = async function (tp, jiraQuery, accountAliases, statusToHeadingMap = {}) {
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

    console.log(`Found ${jiraTasks.length} Jira tasks to process`);

    // Create a map of Jira keys to their current status for quick lookup
    const jiraStatusMap = {};
    jiraTasks.forEach(task => {
      jiraStatusMap[task.key] = task.fields.status.name;
    });

    // First, add new tasks that don't exist in the Kanban
    console.log('Starting to add new tasks to Kanban board');
    const addResult = await addNewTasks(kanbanContent, jiraTasks, statusToHeadingMap, activeFile);
    kanbanContent = addResult.content;

    // Output the number of tasks added
    if (addResult.addedCount > 0) {
      console.log(`Successfully added ${addResult.addedCount} new tasks to the Kanban board`);
    } else {
      console.log('No new tasks were added to the Kanban board');
    }

    // Now, update existing tasks based on their current Jira status
    console.log('Starting to update existing tasks in Kanban board');
    await updateExistingTasks(kanbanContent, jiraStatusMap, statusToHeadingMap, activeFile);
  }

  // New function to update existing tasks in the Kanban board
  async function updateExistingTasks(kanbanContent, jiraStatusMap, statusToHeadingMap, activeFile) {
    // Extract all task references from the Kanban content
    const taskRegex = /- \[[ x]\] \[\[([A-Z]+-\d+)\]\]/g;
    let match;
    const tasksToUpdate = [];

    while ((match = taskRegex.exec(kanbanContent)) !== null) {
      const taskKey = match[1];
      if (jiraStatusMap[taskKey]) {
        tasksToUpdate.push({
          key: taskKey,
          status: jiraStatusMap[taskKey],
          fullMatch: match[0],
          position: match.index
        });
      }
    }

    console.log(`Found ${tasksToUpdate.length} existing tasks to update`);

    // Extract all headers and their positions
    const headers = [];
    const headerRegex = /##\s+(.+)\n/g;
    while ((match = headerRegex.exec(kanbanContent)) !== null) {
      headers.push({
        name: match[1],
        position: match.index,
        length: match[0].length
      });
    }

    // Sort headers by position (descending) to avoid position shifts when modifying content
    headers.sort((a, b) => b.position - a.position);

    // Create sections map based on headers
    const sections = {};
    for (let i = 0; i < headers.length; i++) {
      const currentHeader = headers[i];
      const nextHeader = headers[i + 1];
      const sectionStart = currentHeader.position + currentHeader.length;
      const sectionEnd = nextHeader ? nextHeader.position : kanbanContent.length;

      sections[currentHeader.name] = {
        start: sectionStart,
        end: sectionEnd,
        content: kanbanContent.substring(sectionStart, sectionEnd)
      };
    }

    // Process tasks to update (in reverse order to maintain positions)
    tasksToUpdate.sort((a, b) => b.position - a.position);

    let updatedContent = kanbanContent;
    let tasksUpdated = 0;

    for (const task of tasksToUpdate) {
      // Find which section the task is currently in
      let currentSection = null;
      for (const [headerName, section] of Object.entries(sections)) {
        if (task.position >= section.start && task.position < section.end) {
          currentSection = headerName;
          break;
        }
      }

      if (!currentSection) {
        console.log(`Could not determine current section for task ${task.key}`);
        continue;
      }

      // Determine target section based on Jira status
      const targetHeader = statusToHeadingMap[task.status] || task.status;

      // If task is already in the correct section, skip it
      if (currentSection === targetHeader) {
        console.log(`Task ${task.key} is already in the correct section: ${targetHeader}`);
        continue;
      }

      console.log(`Moving task ${task.key} from "${currentSection}" to "${targetHeader}"`);

      // Remove task from current position
      updatedContent = updatedContent.slice(0, task.position) +
        updatedContent.slice(task.position + task.fullMatch.length);

      // Find the target header position
      const targetHeaderPos = updatedContent.indexOf(`## ${targetHeader}\n`);
      if (targetHeaderPos === -1) {
        // If target header doesn't exist, create it and add the task
        updatedContent += `\n## ${targetHeader}\n${task.fullMatch}\n`;
      } else {
        // Find where to insert in the target section
        const insertPos = targetHeaderPos + `## ${targetHeader}\n`.length;

        // Check if there's a "Complete" marker
        const nextHeaderPos = updatedContent.indexOf('## ', insertPos);
        const sectionEnd = nextHeaderPos !== -1 ? nextHeaderPos : updatedContent.length;
        const sectionContent = updatedContent.substring(insertPos, sectionEnd);
        const completeMarker = '**Complete**';
        const hasCompleteMarker = sectionContent.includes(completeMarker);

        let insertPosition;
        if (hasCompleteMarker) {
          const completePos = insertPos + sectionContent.indexOf(completeMarker);
          insertPosition = completePos + completeMarker.length + 1;
        } else {
          insertPosition = insertPos;
        }

        // Insert task at the correct position
        updatedContent = updatedContent.slice(0, insertPosition) +
          task.fullMatch + '\n' +
          updatedContent.slice(insertPosition);
      }

      tasksUpdated++;
    }

    if (tasksUpdated > 0) {
      console.log(`Updated ${tasksUpdated} tasks in the Kanban board`);
      await app.vault.modify(activeFile, updatedContent);
    } else {
      console.log('No tasks needed to be moved');
    }
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

  // Function to add new Jira tasks to the Kanban board
  async function addNewTasks(kanbanContent, jiraTasks, statusToHeadingMap, activeFile) {
    console.log('Starting to add new tasks to Kanban board');

    // Extract existing task keys from the Kanban content
    const existingTaskRegex = /\[\[([A-Z]+-\d+)\]\]/g;
    const existingTaskKeys = new Set();
    let match;

    while ((match = existingTaskRegex.exec(kanbanContent)) !== null) {
      existingTaskKeys.add(match[1]);
    }

    console.log(`Found ${existingTaskKeys.size} existing tasks in Kanban`);

    // Identify new tasks that don't exist in the Kanban yet
    const newTasks = jiraTasks.filter(task => !existingTaskKeys.has(task.key));
    console.log(`Found ${newTasks.length} new tasks to add to Kanban`);

    if (newTasks.length === 0) {
      console.log('No new tasks to add');
      return kanbanContent;
    }

    let updatedContent = kanbanContent;
    let tasksAdded = 0;

    // Add each new task to the appropriate section
    for (const task of newTasks) {
      const status = task.fields.status.name;
      const targetHeader = statusToHeadingMap[status] || status;
      const taskEntry = `- [ ] [[${task.key}]] ${task.fields.summary}\n`;

      console.log(`Adding task ${task.key} to section "${targetHeader}"`);

      // Find the target header position
      const targetHeaderPos = updatedContent.indexOf(`## ${targetHeader}\n`);

      if (targetHeaderPos === -1) {
        // If target header doesn't exist, create it and add the task
        console.log(`Creating new section "${targetHeader}" for task ${task.key}`);
        updatedContent += `\n## ${targetHeader}\n${taskEntry}`;
      } else {
        // Find where to insert in the target section
        const insertPos = targetHeaderPos + `## ${targetHeader}\n`.length;

        // Insert task at the beginning of the section
        updatedContent = updatedContent.slice(0, insertPos) +
          taskEntry +
          updatedContent.slice(insertPos);
      }

      tasksAdded++;
    }

    console.log(`Added ${tasksAdded} new tasks to the Kanban board`);
    await app.vault.modify(activeFile, updatedContent);
    return { content: updatedContent, addedCount: tasksAdded };
  }

  // Expose functions to Templater
  return {
    updateKanbanWithJiraTasks,
    fetchJiraTasks,
    searchJiraIssues,
    extractHeadersFromContent,
    fileExists,
    addNewTasks
  };
};
