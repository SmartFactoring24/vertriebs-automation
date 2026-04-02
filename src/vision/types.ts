type VisionCaptureRegion = {
  x: number;
  y: number;
  width: number;
  height: number;
};

type VisionCaptureArtifact = {
  imagePath: string;
  region: VisionCaptureRegion;
  createdAt: string;
};

export type { VisionCaptureArtifact, VisionCaptureRegion };
