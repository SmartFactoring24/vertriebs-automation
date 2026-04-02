import { mkdir } from "node:fs/promises";
import path from "node:path";
import type { AppConfig } from "../config.js";
import type { VisionCaptureArtifact, VisionCaptureRegion } from "./types.js";
import { cropImageRegionToPng } from "./tree-capture.js";

type VisionTemplateManifest = {
  treeGruppenAkteHeaderPath: string;
  treeSubmittedUnitsSelectedPath: string;
  contentOpenPathPath: string;
};

function buildTemplateRegion(
  artifact: VisionCaptureArtifact,
  relX: number,
  relY: number,
  relWidth: number,
  relHeight: number
): VisionCaptureRegion {
  return {
    x: Math.round(artifact.region.width * relX),
    y: Math.round(artifact.region.height * relY),
    width: Math.max(1, Math.round(artifact.region.width * relWidth)),
    height: Math.max(1, Math.round(artifact.region.height * relHeight))
  };
}

async function captureVisionTemplates(options: {
  config: AppConfig;
  fullWindowArtifact: VisionCaptureArtifact;
  headerArtifact: VisionCaptureArtifact;
  treeArtifact: VisionCaptureArtifact;
}): Promise<VisionTemplateManifest> {
  const templatesDir = path.join(options.config.visionDirectory, "templates");
  await mkdir(templatesDir, { recursive: true });

  const treeGruppenAkteHeaderPath = path.join(templatesDir, "tree-gruppen-akte-header.png");
  const treeSubmittedUnitsSelectedPath = path.join(templatesDir, "tree-submitted-units-selected.png");
  const contentOpenPathPath = path.join(templatesDir, "content-open-path.png");

  await cropImageRegionToPng(
    options.fullWindowArtifact.imagePath,
    {
      x: 13,
      y: 152,
      width: 268 - 13,
      height: 184 - 152
    },
    treeGruppenAkteHeaderPath
  );
  await cropImageRegionToPng(
    options.treeArtifact.imagePath,
    buildTemplateRegion(options.treeArtifact, 0.02, 0.34, 0.96, 0.10),
    treeSubmittedUnitsSelectedPath
  );
  await cropImageRegionToPng(
    options.fullWindowArtifact.imagePath,
    {
      x: 276,
      y: 117,
      width: 771 - 276,
      height: 140 - 117
    },
    contentOpenPathPath
  );

  return {
    treeGruppenAkteHeaderPath,
    treeSubmittedUnitsSelectedPath,
    contentOpenPathPath
  };
}

export { captureVisionTemplates };
export type { VisionTemplateManifest };
