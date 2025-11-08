import { build, emptyDir } from "@deno/dnt"

const outDir = './build/npm';

await emptyDir(outDir);

let [version] = Deno.args;
if (!version) {
	throw new Error('a version argument is required to build the npm package');
}

await build({
	entryPoints: ['./mod.ts'],
	outDir,
	shims: {
		deno: true,
	},
	test: false,
	typeCheck: false,
	scriptModule: false,
	compilerOptions: {
		target: 'ES2020',
		sourceMap: true,
		strict: true,
		lib: ['dom', 'dom.iterable'],
		useUnknownInCatchVariables: true,
		noImplicitAny: true,
	},
	importMap: "imports.json",
	package: {
		// package.json properties
		name: '@feecompass/docxml',
		type: 'module',
		version,
		description: 'DOCX Markup Language (forked from @fontoxml/docxml)',
		license: 'MIT',
		engines: {
			node: '>= 14',
		},
		repository: {
			type: 'git',
			url: 'https://github.com/feecompass/docxml.git',
		},
	},
});

await Deno.copyFile('README.md', `${outDir}/README.md`);
await Deno.copyFile('.npmrc', `${outDir}/.npmrc`);
