import { ServiceScope } from '@microsoft/sp-core-library';

export interface IPoliciesViewerProps {
  serviceScope: ServiceScope;
  imageGalleryName: string;
  imagesToDisplay: number;
}