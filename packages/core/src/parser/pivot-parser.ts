import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import type { ParsedPivotTable, PivotCacheField, PivotCacheRecord, PivotDataField } from '../types';

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  isArray: (name) => ['pivotCacheRecords', 'r', 'pivotTableDefinition', 'pivotFields', 'pivotField', 'rowFields', 'field', 'colFields', 'dataFields', 'dataField', 'rowItems', 'colItems', 'i', 'cacheFields', 'cacheField', 'sharedItems', 's', 'Location'].includes(name),
});

/**
 * 从 xlsx 中解析所有数据透视表
 */
export async function parsePivotTables(buffer: ArrayBuffer): Promise<ParsedPivotTable[]> {
  const tables: ParsedPivotTable[] = [];

  try {
    const zip = await JSZip.loadAsync(buffer);

    // 1. 获取 workbook 中的 pivotCache 关系
    const wbRels = await zip.file('xl/_rels/workbook.xml.rels')?.async('text');
    if (!wbRels) return tables;

    const wbRelsXml = xmlParser.parse(wbRels);
    const wbRelsList = wbRelsXml['Relationships']?.['Relationship'] || [];
    const rels = Array.isArray(wbRelsList) ? wbRelsList : [wbRelsList];

    // 查找 workbook 级别的 pivotCache 关系
    const cacheRel = rels.find((r: any) => r?.['@_Type']?.includes('pivotCache'));
    if (!cacheRel) return tables;

    // 2. 读取所有 worksheet 的关系文件找 pivotTable 引用
    const sheetDrawingMap: Map<number, string[]> = new Map();
    for (let i = 0; i < 100; i++) {
      const relsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
      const content = await zip.file(relsPath)?.async('text');
      if (!content) continue;

      const relsXml = xmlParser.parse(content);
      const relsList = relsXml['Relationships']?.['Relationship'] || [];
      const sheetRels = Array.isArray(relsList) ? relsList : [relsList];

      const pivotRefs: string[] = [];
      for (const rel of sheetRels) {
        if (rel?.['@_Type']?.includes('pivotTable')) {
          pivotRefs.push(rel['@_Target']);
        }
      }
      if (pivotRefs.length > 0) {
        sheetDrawingMap.set(i, pivotRefs);
      }
    }

    if (sheetDrawingMap.size === 0) return tables;

    // 3. 解析每个 pivotTable 定义和 cache
    for (const [sheetIndex, pivotRefs] of sheetDrawingMap.entries()) {
      for (const pivotRef of pivotRefs) {
        const pivotPath = pivotRef.startsWith('../')
          ? `xl/${pivotRef.substring(3)}`
          : pivotRef;

        const pivotContent = await zip.file(pivotPath)?.async('text');
        if (!pivotContent) continue;

        const pivotXml = xmlParser.parse(pivotContent);
        const ptDef = pivotXml['pivotTableDefinition'];

        // 获取 pivotTable 的关系找 cacheDefinition
        const pivotDir = pivotPath.substring(0, pivotPath.lastIndexOf('/'));
        const pivotRelsPath = `${pivotDir}/_rels/${pivotPath.substring(pivotPath.lastIndexOf('/') + 1)}.rels`;
        const pivotRelsContent = await zip.file(pivotRelsPath)?.async('text');
        if (!pivotRelsContent) continue;

        const pivotRelsXml = xmlParser.parse(pivotRelsContent);
        const pivotRelsList = pivotRelsXml['Relationships']?.['Relationship'] || [];
        const pRels = Array.isArray(pivotRelsList) ? pivotRelsList : [pivotRelsList];

        const cacheRelEntry = pRels.find((r: any) => r?.['@_Type']?.includes('pivotCacheDefinition'));
        if (!cacheRelEntry) continue;

        // 4. 解析 cacheDefinition
        const cacheTarget = cacheRelEntry['@_Target'];
        const cacheDefPath = cacheTarget.startsWith('../')
          ? `xl/${cacheTarget.substring(3)}`
          : `${pivotDir}/${cacheTarget}`;

        const cacheDefContent = await zip.file(cacheDefPath)?.async('text');
        if (!cacheDefContent) continue;

        const cacheDefXml = xmlParser.parse(cacheDefContent);
        const pcd = cacheDefXml['pivotCacheDefinition'];
        const cacheFieldsNodes = pcd?.['cacheFields'] || [];
        const cacheFieldsList = Array.isArray(cacheFieldsNodes) ? cacheFieldsNodes : [cacheFieldsNodes];

        const cacheFields: PivotCacheField[] = [];
        for (const cf of cacheFieldsList) {
          const cfList = Array.isArray(cf['cacheField']) ? cf['cacheField'] : [cf['cacheField']].filter(Boolean);
          for (const field of cfList) {
            const name = field['@_name'] || '';
            const sharedItemsNode = field['sharedItems']?.[0];
            let sharedItems: string[] = [];
            let isNumeric = true;

            if (sharedItemsNode) {
              const sList = sharedItemsNode['s'] || [];
              const items = Array.isArray(sList) ? sList : [sList].filter(Boolean);
              sharedItems = items.map((s: any) => {
                const v = s['@_v'];
                if (v !== undefined) return String(v);
                const cap = s['@_c'];
                if (cap !== undefined) return String(cap);
                return '';
              });
              // If there are shared items, it's likely text
              if (items.length > 0) isNumeric = false;
              // Check if items have numeric values in shared items
              if (items.length === 0 && sharedItemsNode['@_containsNum'] === '1') {
                isNumeric = true;
              }
            }

            cacheFields.push({ name, sharedItems, isNumeric });
          }
        }

        // 5. 通过关系找到 cacheRecords
        const cacheDefDir = cacheDefPath.substring(0, cacheDefPath.lastIndexOf('/'));
        const cacheDefRelsPath = `${cacheDefDir}/_rels/${cacheDefPath.substring(cacheDefPath.lastIndexOf('/') + 1)}.rels`;
        const cacheDefRelsContent = await zip.file(cacheDefRelsPath)?.async('text');
        let cacheRecordsContent: string | undefined;

        if (cacheDefRelsContent) {
          const cdRelsXml = xmlParser.parse(cacheDefRelsContent);
          const cdRelsList = cdRelsXml['Relationships']?.['Relationship'] || [];
          const cdRels = Array.isArray(cdRelsList) ? cdRelsList : [cdRelsList];
          const recordsRel = cdRels.find((r: any) => r?.['@_Type']?.includes('pivotCacheRecords'));
          if (recordsRel) {
            const recordsTarget = recordsRel['@_Target'];
            const recordsPath = recordsTarget.startsWith('../')
              ? `xl/${recordsTarget.substring(3)}`
              : `${cacheDefDir}/${recordsTarget}`;
            cacheRecordsContent = await zip.file(recordsPath)?.async('text');
          }
        }

        // 6. 解析 cacheRecords
        const cacheRecords: PivotCacheRecord[] = [];
        if (cacheRecordsContent) {
          const recordsXml = xmlParser.parse(cacheRecordsContent);
          const pcr = recordsXml['pivotCacheRecords'];
          const rNodes = pcr?.['r'] || [];
          const records = Array.isArray(rNodes) ? rNodes : [rNodes].filter(Boolean);

          for (const r of records) {
            const values: (string | number | null)[] = [];
            const keys = Object.keys(r).filter(k => k !== '@__prefix');
            for (const key of keys) {
              const v = r[key];
              if (key === '@_') continue;

              if (key === 'x' || key === '@_x') {
                // Shared item index - figure out the field index from position
                let fieldIdx = values.length;
                const sharedItems = cacheFields[fieldIdx]?.sharedItems || [];
                const idx = parseInt(v?.['@_v'] ?? v ?? '0', 10);
                values.push(sharedItems[idx] || null);
              } else if (v?.['@_v'] !== undefined) {
                const raw = v['@_v'];
                const num = parseFloat(raw);
                values.push(isNaN(num) ? raw : num);
              } else {
                values.push(null);
              }
            }
            cacheRecords.push({ values });
          }
        }

        // If no cache records found, try to parse inline from cache definition
        if (cacheRecords.length === 0) {
          const recordsAttr = pcd?.['@_recordCount'];
          const count = parseInt(recordsAttr || '0', 10);
          if (count > 0) {
            // Try to read inline data from cacheFields sharedItems
            const maxRecords = Math.max(1, count);
            for (let i = 0; i < maxRecords; i++) {
              const values: (string | number | null)[] = [];
              for (const field of cacheFields) {
                values.push(field.sharedItems[i] || null);
              }
              cacheRecords.push({ values });
            }
          }
        }

        // 7. 提取布局信息
        const rowFieldIndices: number[] = [];
        const colFieldIndices: number[] = [];
        const dataFields: PivotDataField[] = [];

        const rowFieldsNode = ptDef?.['rowFields']?.[0];
        if (rowFieldsNode) {
          const fieldList = rowFieldsNode['field'] || [];
          const fields = Array.isArray(fieldList) ? fieldList : [fieldList].filter(Boolean);
          for (const f of fields) {
            const idx = parseInt(f['@_x'] ?? f['@_v'] ?? '0', 10);
            rowFieldIndices.push(idx);
          }
        }

        const colFieldsNode = ptDef?.['colFields']?.[0];
        if (colFieldsNode) {
          const fieldList = colFieldsNode['field'] || [];
          const fields = Array.isArray(fieldList) ? fieldList : [fieldList].filter(Boolean);
          for (const f of fields) {
            const idx = parseInt(f['@_x'] ?? f['@_v'] ?? '0', 10);
            colFieldIndices.push(idx);
          }
        }

        const dataFieldsNode = ptDef?.['dataFields']?.[0];
        if (dataFieldsNode) {
          const dfList = dataFieldsNode['dataField'] || [];
          const dfs = Array.isArray(dfList) ? dfList : [dfList].filter(Boolean);
          for (const df of dfs) {
            const fld = parseInt(df['@_fld'] ?? '0', 10);
            const name = df['@_name'] || cacheFields[fld]?.name || `Field ${fld}`;
            const sf = df['@_summarizeFunction'] || 'sum';
            dataFields.push({
              fieldIndex: fld,
              name,
              summarizeFunction: sf as any,
            });
          }
        }

        const tableName = ptDef?.['@_name'] || `PivotTable${tables.length + 1}`;

        tables.push({
          name: tableName,
          sheetIndex,
          cacheFields,
          cacheRecords,
          rowFieldIndices,
          colFieldIndices,
          dataFields,
        });
      }
    }
  } catch (err) {
    console.warn('[excel-preview] Failed to parse pivot tables:', err);
  }

  return tables;
}
