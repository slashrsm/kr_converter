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
                contents = contents.replace(placeholder, `var version = '"${(new Date()).toISOString()}"';`);

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
