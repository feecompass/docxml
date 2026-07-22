import { expect } from 'std/expect';
import { describe, it } from 'std/testing/bdd';

import { create } from '../../utilities/src/dom.ts';
import { twip } from '../../utilities/src/length.ts';
import { ALL_NAMESPACE_DECLARATIONS } from '../../utilities/src/namespaces.ts';
import {
	createObjectRoundRobinTest,
	createXmlRoundRobinTest,
} from '../../utilities/src/tests.ts';
import {
	type SectionProperties,
	sectionPropertiesFromNode,
	sectionPropertiesToNode,
} from '../src/section-properties.ts';

const test = createXmlRoundRobinTest<SectionProperties>(
	sectionPropertiesFromNode,
	sectionPropertiesToNode,
);

const reverseTest = createObjectRoundRobinTest<SectionProperties>(
	sectionPropertiesToNode,
	sectionPropertiesFromNode,
);

const date = new Date();

describe('Section formatting', () => {
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:pgSz
				w:w="1200"
				w:h="1600"
				w:orient="landscape"
			/>
			<w:pgMar
				w:top="1000"
				w:right="1000"
				w:bottom="1000"
				w:left="1000"
				w:header="1000"
				w:footer="1000"
				w:gutter="1000"
			/>
		</w:sectPr>`,
		{
			pageWidth: twip(1200),
			pageHeight: twip(1600),
			pageOrientation: 'landscape',
			pageMargin: {
				top: twip(1000),
				right: twip(1000),
				bottom: twip(1000),
				left: twip(1000),
				header: twip(1000),
				footer: twip(1000),
				gutter: twip(1000),
			},
		},
	);
});

describe('Section type', () => {
	reverseTest(
		{ sectionType: 'continuous' },
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:type w:val="continuous"/>
		</w:sectPr>`,
	);

	reverseTest(
		{ sectionType: 'nextPage' },
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:type w:val="nextPage"/>
		</w:sectPr>`,
	);

	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:type w:val="continuous"/>
		</w:sectPr>`,
		{ sectionType: 'continuous' },
	);
});

describe('Section property change', () => {
	// Node with id, author and date
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:pgSz w:orient="portrait"/>
			<w:sectPrChange w:id="0" w:author="Gabe" w:date="${date.toISOString()}">
				<w:sectPr>
					<w:pgSz w:orient="portrait" w:w="12240" w:h="15840" /> 
				</w:sectPr>
			</w:sectPrChange> 
		</w:sectPr>`,
		{
			pageOrientation: 'portrait',
			change: {
				id: 0,
				author: 'Gabe',
				date: date,
				pageWidth: twip(12240),
				pageHeight: twip(15840),
			},
		},
	);

	// Node with id and date, but without author
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
				<w:pgSz w:orient="portrait"/>
				<w:sectPrChange w:id="0" w:date="${date.toISOString()}">
					<w:sectPr>
						<w:pgSz w:orient="portrait" w:w="12240" w:h="15840" /> 
					</w:sectPr>
				</w:sectPrChange> 
			</w:sectPr>`,
		{
			pageOrientation: 'portrait',
			change: {
				id: 0,
				date: date,
				pageWidth: twip(12240),
				pageHeight: twip(15840),
			},
		},
	);
	// Node with id and author, but without date
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:pgSz w:orient="portrait"/>
			<w:sectPrChange w:id="0" w:author="Gabe">
				<w:sectPr>
					<w:pgSz w:orient="portrait" w:w="12240" w:h="15840" /> 
				</w:sectPr>
			</w:sectPrChange> 
		</w:sectPr>`,
		{
			pageOrientation: 'portrait',
			change: {
				id: 0,
				author: 'Gabe',
				pageWidth: twip(12240),
				pageHeight: twip(15840),
			},
		},
	);

	// Node with id, but without date and author
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
				<w:pgSz w:orient="portrait"/>
				<w:sectPrChange w:id="0">
					<w:sectPr>
						<w:pgSz w:orient="portrait" w:w="12240" w:h="15840" /> 
					</w:sectPr>
				</w:sectPrChange> 
			</w:sectPr>`,
		{
			pageOrientation: 'portrait',
			change: {
				id: 0,
				pageWidth: twip(12240),
				pageHeight: twip(15840),
			},
		},
	);
});

describe('Section column formatting for equally sized columns', () => {
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:cols w:num="3" w:equalWidth="1" w:sep="0" w:space="720"/>
		</w:sectPr>`,
		{
			columns: {
				numberOfColumns: 3,
				equalWidth: true,
				separator: false,
				columnSpace: twip(720),
				columnDefs: [],
			},
		},
	);
});

describe('Section column formatting for differently sized columns', () => {
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:cols w:num="3" w:equalWidth="0" w:sep="1" w:space="720" >
				<w:col w:w="1440" w:space="720"/>
				<w:col w:w="1440" w:space="720" />
				<w:col w:w="2880" />
			</w:cols>
		</w:sectPr>`,
		{
			columns: {
				numberOfColumns: 3,
				equalWidth: false,
				separator: true,
				columnSpace: twip(720),
				columnDefs: [
					{ columnWidth: twip(1440), columnSpace: twip(720) },
					{ columnWidth: twip(1440), columnSpace: twip(720) },
					{ columnWidth: twip(2880) },
				],
			},
		},
	);
});

describe('Section column formatting for with missing properties', () => {
	reverseTest(
		{
			columns: {
				numberOfColumns: 3,
				equalWidth: true,
			},
		},
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:cols w:num="3" w:equalWidth="1" />
		</w:sectPr>`,
	);

	reverseTest(
		{
			columns: {
				columnDefs: [
					{ columnWidth: twip(1440), columnSpace: twip(720) },
					{ columnWidth: twip(1440) },
				],
			},
		},
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:cols w:num="2" w:equalWidth="0">
				<w:col w:w="1440" w:space="720" />
				<w:col w:w="1440"/>
			</w:cols>
		</w:sectPr>`,
	);
});

describe('Section header/footer references', () => {
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:headerReference r:id="test1" w:type="default" />
			<w:headerReference r:id="test2" w:type="first" />
			<w:headerReference r:id="test3" w:type="even" />
		</w:sectPr>`,
		{
			headers: {
				first: 'test2',
				even: 'test3',
				odd: 'test1',
			},
			footers: {
				first: null,
				even: null,
				odd: null,
			},
		},
	);
});

describe('Section titlePg', () => {
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
		</w:sectPr>`,
		{
			isTitlePage: false,
		},
	);
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:titlePg />
		</w:sectPr>`,
		{
			isTitlePage: true,
		},
	);
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:titlePg w:val="1" />
		</w:sectPr>`,
		{
			isTitlePage: true,
		},
	);
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:titlePg w:val="0" />
		</w:sectPr>`,
		{
			isTitlePage: false,
		},
	);
});

describe('Section page and line numbering', () => {
	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:pgNumType w:start="1"/>
		</w:sectPr>`,
		{
			pageNumbering: { start: 1 },
		},
	);

	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:pgNumType/>
		</w:sectPr>`,
		{
			pageNumbering: {},
		},
	);

	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:pgNumType
				w:start="3"
				w:fmt="upperRoman"
				w:chapStyle="2"
				w:chapSep="colon"
			/>
		</w:sectPr>`,
		{
			pageNumbering: {
				start: 3,
				format: 'upperRoman',
				chapterStyle: 2,
				chapterSeparator: 'colon',
			},
		},
	);

	it('omits format when w:fmt attribute is absent', () => {
		const result = sectionPropertiesFromNode(
			create(
				`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
					<w:pgNumType w:start="1"/>
				</w:sectPr>`,
			),
		);
		expect(result.pageNumbering?.start).toBe(1);
		expect('format' in (result.pageNumbering || {})).toBe(false);
	});

	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:lnNumType w:countBy="5"/>
		</w:sectPr>`,
		{
			lineNumbering: { countBy: 5 },
		},
	);

	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:lnNumType
				w:countBy="5"
				w:start="10"
				w:distance="360"
				w:restart="continuous"
			/>
		</w:sectPr>`,
		{
			lineNumbering: {
				countBy: 5,
				start: 10,
				distance: twip(360),
				restart: 'continuous',
			},
		},
	);

	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:lnNumType w:countBy="5"/>
			<w:pgNumType w:start="1"/>
		</w:sectPr>`,
		{
			lineNumbering: { countBy: 5 },
			pageNumbering: { start: 1 },
		},
	);

	reverseTest(
		{ pageNumbering: { start: 1 } },
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:pgNumType w:start="1"/>
		</w:sectPr>`,
	);

	reverseTest(
		{ lineNumbering: { countBy: 5 } },
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:lnNumType w:countBy="5"/>
		</w:sectPr>`,
	);

	reverseTest(
		{ pageNumbering: null },
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
		</w:sectPr>`,
	);

	reverseTest(
		{ pageNumbering: {} },
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:pgNumType/>
		</w:sectPr>`,
	);

	test(
		`<w:sectPr ${ALL_NAMESPACE_DECLARATIONS}>
			<w:pgSz w:orient="portrait"/>
			<w:sectPrChange w:id="0" w:author="Gabe" w:date="${date.toISOString()}">
				<w:sectPr>
					<w:pgNumType w:start="1"/>
				</w:sectPr>
			</w:sectPrChange>
		</w:sectPr>`,
		{
			pageOrientation: 'portrait',
			change: {
				id: 0,
				author: 'Gabe',
				date: date,
				pageNumbering: { start: 1 },
			},
		},
	);
});
