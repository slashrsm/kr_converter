const result = await Bun.build({
  entrypoints: ["./index.html"],
  outdir: "./docs",
  target: "browser",
  plugins: [
    {
      name: "Application version string",
      setup(build) {
        const placeholder = "var version = 'development_version';";

        build.onLoad({ filter: /script\.js/ }, async (args) => {
            try {
                // Read the HTML file content
                let contents = await Bun.file(args.path).text();

                // Check for placeholder
                if (!contents.includes(placeholder)) {
                    throw new Error(`Version placeholder not found in ${args.path}`);
                }

                // Replace placeholder with version tag
                const date = new Date();
                const year = date.getFullYear();
                const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are 0-based
                const day = String(date.getDate()).padStart(2, '0');
                const hours = String(date.getHours()).padStart(2, '0');
                const minutes = String(date.getMinutes()).padStart(2, '0');
                const ts_string = `${year}-${month}-${day} ${hours}:${minutes}`;
                contents = contents.replace(placeholder, `var version = '${ts_string}';`);

                // Return the modified content
                return {
                    contents,
                    loader: 'js',
                };
            } catch (error) {
                console.error('Error in version-inject plugin:', error.message);
                throw error; // Fail the build if there's an error
            }
        });
      },
    }
  ],
});

if (!result.success) {
  console.error("Build failed:", result.logs);
  process.exit(1);
}

console.log("Build succeeded");
