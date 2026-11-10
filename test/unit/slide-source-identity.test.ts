import { expect, it } from 'vitest';
import { importedSlideShapeId, installSlideSourceIdentity } from '../../lib/onlyoffice/guards/slide-source-identity';

it('captures package IDs before the native loader renumbers them without changing the loader', () => {
  class Loader {
    ReadNvUniProp(shape: object) {
      return { cNvPr: { id: Reflect.get(shape, 'packageId') } };
    }
  }
  const scope = { AscCommon: { BinaryPPTYLoader: Loader } };
  const win = scope as unknown as Window;
  expect(installSlideSourceIdentity(win)).toBe(true);
  const wrapped = scope.AscCommon.BinaryPPTYLoader;
  expect(installSlideSourceIdentity(win)).toBe(true);
  expect(scope.AscCommon.BinaryPPTYLoader).toBe(wrapped);
  const loader = new wrapped();
  expect(loader).toBeInstanceOf(Loader);
  const shape = { packageId: 42 };
  const props = loader.ReadNvUniProp(shape);
  props.cNvPr.id = 987;
  expect(importedSlideShapeId(shape)).toBe('42');
  expect(importedSlideShapeId({ ...shape })).toBeUndefined();
  const invalid = { packageId: 'invalid' };
  loader.ReadNvUniProp(invalid);
  expect(importedSlideShapeId(invalid)).toBeUndefined();
  expect(installSlideSourceIdentity({} as Window)).toBe(false);
});

it('captures the first import when the full SDK constructor arrives after app-ready', () => {
  const scope: {
    AscCommon: { BinaryPPTYLoader?: new () => { ReadNvUniProp(shape: object): { cNvPr: { id: number } } } };
  } = { AscCommon: {} };
  expect(installSlideSourceIdentity(scope as unknown as Window)).toBe(true);
  scope.AscCommon.BinaryPPTYLoader = class {
    ReadNvUniProp() {
      return { cNvPr: { id: 7 } };
    }
  };
  const shape = {};
  new scope.AscCommon.BinaryPPTYLoader().ReadNvUniProp(shape);
  expect(importedSlideShapeId(shape)).toBe('7');
});
