import { useEffect, useState } from 'react';
import { db } from '@/db/db';

/*
 * Avatar rule (docs/01): never cropped or distorted — object-fit: contain,
 * native aspect ratio, rounded corners ok, no circle-crop.
 */
export default function Avatar({
  assetId,
  name,
  className = 'w-12 h-12',
}: {
  assetId?: string;
  name: string;
  className?: string;
}) {
  const url = useAssetUrl(assetId);
  if (!url) {
    return (
      <div
        className={`${className} shrink-0 rounded-xl bg-surface-2 grid place-items-center text-muted font-medium`}
        aria-hidden
      >
        {name.slice(0, 1).toUpperCase() || '?'}
      </div>
    );
  }
  return (
    <img
      src={url}
      alt={name}
      className={`${className} shrink-0 rounded-xl object-contain bg-surface-2`}
    />
  );
}

export function useAssetUrl(assetId?: string): string | undefined {
  const [url, setUrl] = useState<string>();
  useEffect(() => {
    if (!assetId) {
      setUrl(undefined);
      return;
    }
    let objectUrl: string | undefined;
    let cancelled = false;
    db.galleryAssets.get(assetId).then((asset) => {
      if (asset && !cancelled) {
        objectUrl = URL.createObjectURL(asset.blob);
        setUrl(objectUrl);
      }
    });
    return () => {
      cancelled = true;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [assetId]);
  return url;
}
