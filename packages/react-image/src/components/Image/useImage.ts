import { useImageAria } from '@fluentui/accessibility';
import { ComposePreparedOptions } from '@fluentui/react-compose';
import { getStyleFromPropsAndOptions } from '@fluentui/react-theme-provider';
/* eslint-disable @typescript-eslint/ban-ts-comment */
// @ts-ignore Typings require esModuleInterop
import objectFitImages from 'object-fit-images';
/* eslint-enable @typescript-eslint/ban-ts-comment */
import * as React from 'react';

import { ImageProps, ImageState } from './Image.types';

const isFitSupported = (function() {
  const testImg = new Image();

  return 'object-fit' in testImg.style && 'object-position' in testImg.style;
})();

/**
 * The useImage hook processes the Image component props and returns state.
 */
export const useImage = (
  props: ImageProps,
  ref: React.Ref<HTMLElement>,
  options: ComposePreparedOptions,
): ImageState => {
  const imageRef = React.useRef<HTMLElement>(null);
  const ariaProps = useImageAria(props);

  const handleImageLoad: React.DOMAttributes<HTMLImageElement>['onLoad'] = e => {
    if (props.onLoad) {
      props.onLoad(e);
    }

    if (imageRef.current) {
      const rect = imageRef.current.getBoundingClientRect();
      let actualImageSize: 'small' | 'medium' | 'large' = 'small';

      if (rect.width > 80) {
        actualImageSize = 'large';
      } else if (rect.width > 32) {
        actualImageSize = 'medium';
      }

      imageRef.current.setAttribute('data-image-size', actualImageSize);
    }
  };

  React.useEffect(() => {
    if (!isFitSupported) {
      objectFitImages(imageRef.current);
    }
  }, []);

  return {
    ...ariaProps,
    ...props,
    imageRef,
    onLoad: handleImageLoad,
    style: getStyleFromPropsAndOptions(props, options, '--image'),
  };
};
