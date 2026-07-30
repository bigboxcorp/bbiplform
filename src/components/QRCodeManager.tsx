import React, { useState, useEffect } from 'react';
import { QRCodeCanvas } from 'qrcode.react';
import { Link2, Image as ImageIcon, Video, Plus, Trash2, Edit, X, Save, Copy, ExternalLink, QrCode, Download } from 'lucide-react';

interface QRCodeData {
  id: string;
  title: string;
  type: 'link' | 'image' | 'video';
  targetData: string;
  fgColor?: string;
  bgColor?: string;
  logoUrl?: string;
  createdAt: string;
  updatedAt: string;
}

export default function QRCodeManager() {
  const [qrs, setQrs] = useState<QRCodeData[]>([]);
  const [loading, setLoading] = useState(true);
  const [isEditing, setIsEditing] = useState<string | null>(null);
  const [isCreating, setIsCreating] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [viewingQr, setViewingQr] = useState<QRCodeData | null>(null);
  
  const [editForm, setEditForm] = useState<Partial<QRCodeData>>({});

  useEffect(() => {
    fetchQrs();
  }, []);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (editForm.type === 'image' && file.size > 20 * 1024 * 1024) {
      alert("Image size should be less than 20MB.");
      return;
    }
    if (editForm.type === 'video' && file.size > 200 * 1024 * 1024) {
      alert("Video size should be less than 200MB.");
      return;
    }
    if (editForm.type === 'link' && file.size > 100 * 1024 * 1024) {
      alert("Document size should be less than 100MB.");
      return;
    }

    const formData = new FormData();
    formData.append('file', file);
    
    setIsUploading(true);
    try {
      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (res.ok) {
        setEditForm(prev => ({ ...prev, targetData: data.url }));
      } else {
        alert(data.error || "Failed to upload file");
      }
    } catch (err) {
      alert("Error uploading file");
    } finally {
      setIsUploading(false);
    }
  };

  const handleLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 5 * 1024 * 1024) {
      alert("Logo size should be less than 5MB.");
      return;
    }

    const formData = new FormData();
    formData.append('file', file);
    
    setIsUploading(true);
    try {
      const res = await fetch('/api/upload', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (res.ok) {
        setEditForm(prev => ({ ...prev, logoUrl: data.url }));
      } else {
        alert(data.error || "Failed to upload logo");
      }
    } catch (err) {
      alert("Error uploading logo");
    } finally {
      setIsUploading(false);
    }
  };

  const fetchQrs = () => {
    setLoading(true);
    fetch('/api/qrcodes')
      .then(res => res.json())
      .then(data => {
        setQrs(data);
        setLoading(false);
      })
      .catch(err => {
        console.error(err);
        setLoading(false);
      });
  };

  const safeOrigin = window.location.hostname.includes('localhost') || window.location.hostname.includes('127.0.0.1') 
    ? window.location.origin 
    : window.location.origin.replace(/^http:\/\//i, 'https://');

  const handleCreateNew = () => {
    setIsCreating(true);
    setEditForm({ title: '', type: 'link', targetData: '', fgColor: '#000000', bgColor: '#FFFFFF', logoUrl: '' });
  };

  const handleSave = () => {
    if (!editForm.title || !editForm.targetData) {
      alert("Please provide both a title and target data (URL).");
      return;
    }
    
    if (!editForm.targetData.startsWith('http://') && !editForm.targetData.startsWith('https://')) {
      alert("Target data must be a valid URL starting with http:// or https://");
      return;
    }

    const url = isCreating ? '/api/qrcodes' : `/api/qrcodes/${isEditing}`;
    const method = isCreating ? 'POST' : 'PUT';

    fetch(url, {
      method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(editForm)
    })
      .then(res => res.json())
      .then(data => {
        setIsCreating(false);
        setIsEditing(null);
        fetchQrs();
      })
      .catch(err => alert("Failed to save QR code"));
  };

  const handleDelete = (id: string) => {
    if (window.confirm("Are you sure you want to delete this dynamic QR code? It will stop working immediately.")) {
      fetch(`/api/qrcodes/${id}`, { method: 'DELETE' })
        .then(() => fetchQrs())
        .catch(() => alert("Failed to delete QR code"));
    }
  };

  const handleDownload = (id: string, title: string) => {
    const canvas = document.getElementById(`qr-canvas-${id}`) as HTMLCanvasElement;
    if (!canvas) return;
    const pngUrl = canvas
      .toDataURL("image/png")
      .replace("image/png", "image/octet-stream");
    let downloadLink = document.createElement("a");
    downloadLink.href = pngUrl;
    downloadLink.download = `${title.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_qr.png`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
  };

  const getTypeIcon = (type: string) => {
    switch (type) {
      case 'link': return <Link2 size={16} />;
      case 'image': return <ImageIcon size={16} />;
      case 'video': return <Video size={16} />;
      default: return <Link2 size={16} />;
    }
  };

  return (
    <div className="flex flex-col gap-6 max-w-6xl w-full mx-auto p-4 sm:p-6 lg:p-8">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center bg-white p-6 rounded-xl border border-slate-200 shadow-sm gap-4">
        <div>
          <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
            <QrCode className="text-purple-600" /> Dynamic QR Code Manager
          </h2>
          <p className="text-sm text-slate-500 mt-1">Generate and manage QR codes that can be updated anytime without changing the QR image.</p>
        </div>
        
        <div className="flex flex-col sm:items-end gap-3">
          <button 
            onClick={handleCreateNew} 
            disabled={isCreating || isEditing !== null}
            className="bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-lg font-bold text-sm shadow-sm flex items-center justify-center gap-2 transition-all cursor-pointer disabled:opacity-50"
          >
            <Plus size={16} /> Create New QR
          </button>
        </div>
      </div>

      {(isCreating || isEditing) && (
        <div className="bg-white p-6 rounded-xl border border-purple-200 shadow-md">
          <div className="flex justify-between items-center mb-4">
            <h3 className="font-bold text-lg text-slate-800">{isCreating ? 'Create Dynamic QR' : 'Edit Dynamic QR'}</h3>
            <button onClick={() => { setIsCreating(false); setIsEditing(null); }} className="text-slate-400 hover:text-slate-700">
              <X size={20} />
            </button>
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1">Title</label>
              <input 
                type="text" 
                value={editForm.title || ''} 
                onChange={e => setEditForm({ ...editForm, title: e.target.value })}
                placeholder="e.g. My Website, Promo Video..."
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:border-purple-500 focus:ring-1 focus:ring-purple-500 text-sm"
              />
            </div>
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1">Type</label>
              <select 
                value={editForm.type || 'link'}
                onChange={e => setEditForm({ ...editForm, type: e.target.value as any })}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:border-purple-500 focus:ring-1 focus:ring-purple-500 text-sm bg-white"
              >
                <option value="link">Website Link / Document (PDF)</option>
                <option value="image">Image</option>
                <option value="video">Video (MP4)</option>
              </select>
            </div>
          </div>
          
          <div className="mb-4">
            <label className="block text-sm font-semibold text-slate-700 mb-1">Target URL or Upload Media</label>
            <div className="flex gap-2">
              <input 
                type="text" 
                value={editForm.targetData || ''} 
                onChange={e => setEditForm({ ...editForm, targetData: e.target.value })}
                placeholder={editForm.type === 'link' ? "https://example.com" : "Enter URL or Upload File"}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:border-purple-500 focus:ring-1 focus:ring-purple-500 text-sm font-mono"
              />
              <div className="relative overflow-hidden bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg flex items-center justify-center shrink-0 px-3 cursor-pointer transition-colors">
                <span className="text-sm font-semibold text-slate-700">{isUploading ? 'Uploading...' : 'Upload File'}</span>
                <input 
                  type="file" 
                  accept={editForm.type === 'image' ? 'image/*' : editForm.type === 'video' ? 'video/mp4' : '*/*'}
                  onChange={handleFileUpload} 
                  disabled={isUploading}
                  className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed" 
                />
              </div>
            </div>
            <p className="text-xs text-slate-500 mt-2 p-2 bg-slate-50 border border-slate-200 rounded-md">
              <strong className="text-slate-700">Upload Limits:</strong><br/>
              • <strong>Image:</strong> Max 20MB<br/>
              • <strong>Video (MP4):</strong> Max 200MB<br/>
              • <strong>Document/PDF:</strong> Max 100MB<br/>
              <em>You can also paste a direct URL instead of uploading.</em>
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1">QR Color</label>
              <div className="flex gap-2 items-center">
                <input 
                  type="color" 
                  value={editForm.fgColor || '#000000'} 
                  onChange={e => setEditForm({ ...editForm, fgColor: e.target.value })}
                  className="w-8 h-8 rounded cursor-pointer border-0 p-0 bg-transparent"
                />
                <span className="text-xs text-slate-500">{editForm.fgColor || '#000000'}</span>
              </div>
            </div>
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1">Background</label>
              <div className="flex gap-2 items-center">
                <input 
                  type="color" 
                  value={editForm.bgColor || '#ffffff'} 
                  onChange={e => setEditForm({ ...editForm, bgColor: e.target.value })}
                  className="w-8 h-8 rounded cursor-pointer border-0 p-0 bg-transparent shadow-sm border border-slate-200"
                />
                <span className="text-xs text-slate-500">{editForm.bgColor || '#ffffff'}</span>
              </div>
            </div>
            <div>
              <label className="block text-sm font-semibold text-slate-700 mb-1">Logo URL (Optional)</label>
              <div className="flex gap-2">
                <input 
                  type="text" 
                  value={editForm.logoUrl || ''} 
                  onChange={e => setEditForm({ ...editForm, logoUrl: e.target.value })}
                  placeholder="https://.../logo.png"
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:outline-none focus:border-purple-500 text-sm"
                />
                <div className="relative overflow-hidden bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg flex items-center justify-center shrink-0 px-2 cursor-pointer transition-colors" title="Upload Logo (Max 5MB)">
                  <span className="text-xs font-semibold text-slate-700">{isUploading ? '...' : 'Upload'}</span>
                  <input 
                    type="file" 
                    accept="image/*"
                    onChange={handleLogoUpload} 
                    disabled={isUploading}
                    className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed" 
                  />
                </div>
              </div>
            </div>
          </div>

          <div className="flex justify-end gap-3">
            <button 
              onClick={() => { setIsCreating(false); setIsEditing(null); }}
              className="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 font-semibold text-sm hover:bg-slate-50 transition-colors"
            >
              Cancel
            </button>
            <button 
              onClick={handleSave}
              className="px-4 py-2 bg-purple-600 rounded-lg text-white font-semibold text-sm hover:bg-purple-700 transition-colors flex items-center gap-2"
            >
              <Save size={16} /> Save QR
            </button>
          </div>
        </div>
      )}

      {loading ? (
        <div className="text-center p-8 text-slate-500 font-medium animate-pulse">Loading QR codes...</div>
      ) : qrs.length === 0 && !isCreating ? (
        <div className="bg-slate-50 border border-dashed border-slate-300 rounded-xl p-12 text-center flex flex-col items-center">
          <div className="w-16 h-16 bg-white rounded-full flex items-center justify-center shadow-sm mb-4">
            <QrCode className="text-slate-400 w-8 h-8" />
          </div>
          <h3 className="text-lg font-bold text-slate-700">No Dynamic QRs Yet</h3>
          <p className="text-sm text-slate-500 mt-2 mb-6 max-w-sm">Create dynamic QR codes where you can update the destination link anytime without re-printing.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {qrs.map(qr => {
            const qrUrl = `${safeOrigin}/qr/${qr.id}`;
            return (
              <div key={qr.id} className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col group hover:shadow-md hover:border-purple-200 transition-all">
                <div className="p-5 flex gap-4">
                  <div 
                    className="shrink-0 bg-white p-2 rounded-xl border-2 border-slate-100 flex items-center justify-center cursor-pointer hover:border-purple-300 transition-colors"
                    onClick={() => setViewingQr(qr)}
                    title="Click to view larger"
                  >
                    <QRCodeCanvas 
                      id={`qr-canvas-${qr.id}`}
                      value={qrUrl} 
                      size={400} 
                      level="H"
                      style={{ width: 80, height: 80 }}
                      fgColor={qr.fgColor || "#000000"} 
                      bgColor={qr.bgColor || "#ffffff"}
                      imageSettings={qr.logoUrl ? { src: qr.logoUrl, height: 80, width: 80, excavate: true } : undefined}
                    />
                  </div>
                  <div className="flex-1 min-w-0 flex flex-col justify-center">
                    <div className="flex items-center gap-1.5 mb-1 text-xs font-bold text-purple-600 uppercase tracking-wider">
                      {getTypeIcon(qr.type)}
                      {qr.type}
                    </div>
                    <h3 className="font-bold text-slate-800 text-lg truncate mb-1" title={qr.title}>{qr.title}</h3>
                    <p className="text-xs text-slate-500 truncate font-mono" title={qr.targetData}>{qr.targetData}</p>
                  </div>
                </div>
                
                <div className="border-t border-slate-100 p-3 bg-slate-50 flex items-center justify-between group-hover:bg-purple-50/30 transition-colors">
                  <div className="flex items-center gap-2">
                     <button 
                       onClick={() => {
                         navigator.clipboard.writeText(qrUrl);
                         alert('Short link copied!');
                       }}
                       className="text-xs font-bold text-slate-600 hover:text-purple-700 flex items-center gap-1 cursor-pointer bg-white px-2 py-1 rounded shadow-xs border border-slate-200"
                     >
                       <Copy size={12} /> Copy Link
                     </button>
                     <a 
                       href={qrUrl} 
                       target="_blank" 
                       rel="noopener noreferrer"
                       className="text-xs font-bold text-slate-600 hover:text-blue-700 flex items-center gap-1 cursor-pointer bg-white px-2 py-1 rounded shadow-xs border border-slate-200"
                     >
                       <ExternalLink size={12} /> Test
                     </a>
                     <button 
                       onClick={() => handleDownload(qr.id, qr.title)}
                       className="text-xs font-bold text-slate-600 hover:text-green-700 flex items-center gap-1 cursor-pointer bg-white px-2 py-1 rounded shadow-xs border border-slate-200"
                     >
                       <Download size={12} /> Download
                     </button>
                  </div>
                  
                  <div className="flex gap-2">
                    <button 
                      onClick={() => {
                        setIsEditing(qr.id);
                        setEditForm(qr);
                        setIsCreating(false);
                      }}
                      className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 bg-white rounded shadow-xs hover:shadow-sm transition-colors cursor-pointer"
                      title="Edit Target"
                    >
                      <Edit size={14} />
                    </button>
                    <button 
                      onClick={() => handleDelete(qr.id)}
                      className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 bg-white rounded shadow-xs hover:shadow-sm transition-colors cursor-pointer"
                      title="Delete QR"
                    >
                      <Trash2 size={14} />
                    </button>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {viewingQr && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-[100] flex items-center justify-center p-4" onClick={() => setViewingQr(null)}>
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full overflow-hidden" onClick={e => e.stopPropagation()}>
            <div className="p-6 text-center flex flex-col items-center">
              <h3 className="text-lg font-bold text-slate-900 mb-2 truncate w-full" title={viewingQr.title}>{viewingQr.title}</h3>
              <p className="text-xs text-slate-500 mb-4">{viewingQr.targetData}</p>
              
              <div className="bg-white p-4 rounded-xl border-2 border-slate-100 inline-block mb-6">
                <QRCodeCanvas 
                  id={`view-qr-canvas-${viewingQr.id}`}
                  value={`${safeOrigin}/qr/${viewingQr.id}`} 
                  size={240} 
                  level="H"
                  fgColor={viewingQr.fgColor || "#000000"} 
                  bgColor={viewingQr.bgColor || "#ffffff"}
                  imageSettings={viewingQr.logoUrl ? { src: viewingQr.logoUrl, height: 48, width: 48, excavate: true } : undefined}
                />
              </div>
              
              <div className="flex gap-3 w-full">
                <button 
                  onClick={() => setViewingQr(null)}
                  className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-800 font-medium py-2.5 rounded-lg transition-colors cursor-pointer"
                >
                  Close
                </button>
                <button 
                  onClick={() => {
                    const canvas = document.getElementById(`view-qr-canvas-${viewingQr.id}`) as HTMLCanvasElement;
                    if (!canvas) return;
                    const pngUrl = canvas.toDataURL("image/png").replace("image/png", "image/octet-stream");
                    let downloadLink = document.createElement("a");
                    downloadLink.href = pngUrl;
                    downloadLink.download = `${viewingQr.title.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_qr.png`;
                    document.body.appendChild(downloadLink);
                    downloadLink.click();
                    document.body.removeChild(downloadLink);
                  }}
                  className="flex-1 bg-purple-600 hover:bg-purple-700 text-white font-medium py-2.5 rounded-lg transition-colors cursor-pointer flex items-center justify-center gap-2"
                >
                  <Download size={16} /> Download
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
