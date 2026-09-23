# Use cases

I use this to convert my science magazine articles from Microsoft Word (my main editor) to HTML, so Google can index them and show my website in search.

If you're a scientist writing articles, it minimal repo will come in handy. Just install Bun, point it to your word file and done.

To install dependencies:

```bash
bun install
```

To run:

```bash
bun run convert.ts --input="./article1.docx"
```

*it's faster on Linux than Windows.

Writes `public/article1/index.html` and images, using the DOCX filename as the folder name.

Check the CLI with `bun run convert.test.ts`.

This project was created using `bun init` in bun v1.3.14. [Bun](https://bun.com) is a fast all-in-one JavaScript runtime.
