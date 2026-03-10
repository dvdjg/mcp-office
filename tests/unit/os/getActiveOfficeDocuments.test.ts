import getActiveOfficeDocumentsTool from '../../../src/tools/os/getActiveOfficeDocuments.tool';

describe('os/getActiveOfficeDocuments', () => {
  const originalPlatform = process.platform;

  afterEach(() => {
    Object.defineProperty(process, 'platform', { value: originalPlatform, writable: true });
  });

  test('returns an empty list on non-Windows platforms', async () => {
    Object.defineProperty(process, 'platform', { value: 'linux', writable: true });

    const result = await getActiveOfficeDocumentsTool.handler({});

    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.documents).toEqual([]);
    }
  });
});
