// Initialize the add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("fetchRepos").onclick = () => tryCatch(fetchAndInsertRepos);
  }
});

async function fetchAndInsertRepos() {
  const username = document.getElementById("username").value;
  if (!username) {
    alert("Please enter a GitHub username.");
    return;
  }

  try {
    const repos = await fetchGitHubRepos(username);
    await Word.run(async (context) => {
      const docBody = context.document.body;
      docBody.insertText(`Repositories for ${username}:\n`, Word.InsertLocation.end);
      
      repos.forEach(repo => {
        docBody.insertText(`- ${repo.name}\n`, Word.InsertLocation.end);
      });

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function fetchGitHubRepos(username) {
  const response = await fetch(`https://api.github.com/users/${username}/repos`);
  if (!response.ok) {
    throw new Error("Failed to fetch repos");
  }
  return response.json();
}

// Error handling helper
function tryCatch(callback) {
  Promise.resolve()
    .then(callback)
    .catch((error) => {
      console.error(error);
      alert("An error occurred: " + error.message);
    });
}
