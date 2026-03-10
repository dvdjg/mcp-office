import slidesTool from '../../../src/tools/powerpoint/slides.tool';

describe('powerpoint/slides', () => {
  test('library path reports getSlideCount as unsupported', async () => {
    const result = await slidesTool.handler({
      filePath: 'C:/test.pptx',
      operation: 'getSlideCount',
      useComInterop: false,
    });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
    }
  });
});
