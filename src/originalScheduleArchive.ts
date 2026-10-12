import {
  getBytes,
  getMetadata,
  listAll,
  ref as storageRef,
  uploadBytes,
  updateMetadata,
  type StorageReference,
} from 'firebase/storage';
import { storage } from './firebase';

export type OriginalScheduleVersion = {
  path: string;
  name: string;
  uploadedAt: string;
  originalFileName: string;
};

function safeFileName(value: string) {
  return value.replace(/[^a-zA-Z0-9._-]+/g, '_').slice(-140) || 'schedule.xlsx';
}

function weekFolder(weekStartDate: string) {
  return weekStartDate.replace(/[^0-9-]/g, '');
}

/** Keeps every uploaded source workbook as an immutable version in Firebase Storage. */
export async function archiveOriginalSchedule(file: File, weekStartDate: string) {
  if (!weekStartDate) throw new Error('No se puede archivar el Excel sin una fecha inicial de semana.');
  const uploadedAt = new Date().toISOString();
  const versionName = `${uploadedAt.replace(/[:.]/g, '-')}_${safeFileName(file.name)}`;
  const path = `original-schedules/${weekFolder(weekStartDate)}/versions/${versionName}`;
  const fileRef = storageRef(storage, path);
  await uploadBytes(fileRef, file, {
    contentType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    customMetadata: {
      weekStartDate,
      originalFileName: file.name,
      uploadedAt,
      archiveStatus: 'pending',
    },
  });
  return { path, name: versionName, uploadedAt, originalFileName: file.name };
}

/** Returns the most recent original workbook, never a transformed Google Sheets export. */
export async function markOriginalScheduleCommitted(path: string) {
  const fileRef = storageRef(storage, path);
  await updateMetadata(fileRef, { customMetadata: { archiveStatus: 'committed' } });
}

export async function getLatestOriginalSchedule(weekStartDate: string) {
  if (!weekStartDate) return null;
  const folder = storageRef(storage, `original-schedules/${weekFolder(weekStartDate)}/versions`);
  let items: StorageReference[];
  try {
    items = (await listAll(folder)).items;
  } catch (error: any) {
    if (error?.code === 'storage/object-not-found') return null;
    throw error;
  }
  if (!items.length) return null;

  const versions = await Promise.all(items.map(async (item) => {
    const metadata = await getMetadata(item);
    return {
      item,
      uploadedAt: metadata.customMetadata?.uploadedAt || metadata.timeCreated || '',
      originalFileName: metadata.customMetadata?.originalFileName || item.name,
      archiveStatus: metadata.customMetadata?.archiveStatus || 'committed',
    };
  }));
  const committedVersions = versions.filter((version) => version.archiveStatus === 'committed');
  if (!committedVersions.length) return null;
  committedVersions.sort((a, b) => b.uploadedAt.localeCompare(a.uploadedAt));
  const latest = committedVersions[0];
  const bytes = await getBytes(latest.item);
  return {
    bytes,
    path: latest.item.fullPath,
    uploadedAt: latest.uploadedAt,
    originalFileName: latest.originalFileName,
  };
}
