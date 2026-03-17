import React, { useState } from 'react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, ReferenceLine } from 'recharts';
import { TrendingUp, Download, Upload, Plus, Trash2, Edit2, Check, X } from 'lucide-react';

const ProductivityTracker = () => {
  // Pre-loaded historical data
  const initialData = [
    { month: "Jan 2021", productivity: 70.89, sales: 714588, hours: 10070 },
    { month: "Feb 2021", productivity: 74.18, sales: 726072, hours: 9788 },
    { month: "Mar 2021", productivity: 78.79, sales: 820668, hours: 10414 },
    { month: "Apr 2021", productivity: 78.84, sales: 757256, hours: 9604 },
    { month: "May 2021", productivity: 78.85, sales: 881100, hours: 11174 },
    { month: "Jun 2021", productivity: 77.54, sales: 799500, hours: 10312 },
    { month: "Jul 2021", productivity: 73.43, sales: 771204, hours: 10502 },
    { month: "Aug 2021", productivity: 76.06, sales: 841560, hours: 11063 },
    { month: "Sep 2021", productivity: 77.85, sales: 809028, hours: 10393 },
    { month: "Oct 2021", productivity: 76.98, sales: 845064, hours: 10976 },
    { month: "Nov 2021", productivity: 80.05, sales: 817608, hours: 10215 },
    { month: "Dec 2021", productivity: 77.22, sales: 750360, hours: 9716 },
    { month: "Jan 2022", productivity: 78.62, sales: 786420, hours: 10003 },
    { month: "Feb 2022", productivity: 80.80, sales: 814488, hours: 10080 },
    { month: "Mar 2022", productivity: 85.78, sales: 939996, hours: 10957 },
    { month: "Apr 2022", productivity: 88.17, sales: 906936, hours: 10287 },
    { month: "May 2022", productivity: 84.96, sales: 1000200, hours: 11772 },
    { month: "Jun 2022", productivity: 82.73, sales: 924576, hours: 11176 },
    { month: "Jul 2022", productivity: 81.51, sales: 936936, hours: 11494 },
    { month: "Aug 2022", productivity: 84.00, sales: 1012356, hours: 12052 },
    { month: "Sep 2022", productivity: 84.75, sales: 922704, hours: 10887 },
    { month: "Oct 2022", productivity: 86.47, sales: 1014144, hours: 11728 },
    { month: "Nov 2022", productivity: 83.42, sales: 916224, hours: 10983 },
    { month: "Dec 2022", productivity: 84.29, sales: 941772, hours: 11173 },
    { month: "Jan 2023", productivity: 84.38, sales: 976284, hours: 11571 },
    { month: "Feb 2023", productivity: 90.49, sales: 1011732, hours: 11180 },
    { month: "Mar 2023", productivity: 92.78, sales: 1134492, hours: 12230 },
    { month: "Apr 2023", productivity: 94.54, sales: 1071648, hours: 11337 },
    { month: "May 2023", productivity: 93.28, sales: 1190700, hours: 12764 },
    { month: "Jun 2023", productivity: 91.97, sales: 1093080, hours: 11883 },
    { month: "Jul 2023", productivity: 88.73, sales: 1107348, hours: 12482 },
    { month: "Aug 2023", productivity: 88.71, sales: 1112184, hours: 12537 },
    { month: "Sep 2023", productivity: 89.98, sales: 1044780, hours: 11612 },
    { month: "Oct 2023", productivity: 88.31, sales: 1095048, hours: 12399 },
    { month: "Nov 2023", productivity: 86.71, sales: 1015200, hours: 11707 },
    { month: "Dec 2023", productivity: 87.83, sales: 990732, hours: 11278 },
    { month: "Jan 2024", productivity: 85.62, sales: 975528, hours: 11393 },
    { month: "Feb 2024", productivity: 92.06, sales: 1057764, hours: 11490 },
    { month: "Mar 2024", productivity: 93.24, sales: 1118472, hours: 11993 },
    { month: "Apr 2024", productivity: 89.35, sales: 1002468, hours: 11219 },
    { month: "May 2024", productivity: 90.14, sales: 1092324, hours: 12117 },
    { month: "Jun 2024", productivity: 85.88, sales: 1000356, hours: 11648 },
    { month: "Jul 2024", productivity: 82.59, sales: 963948, hours: 11671 },
    { month: "Aug 2024", productivity: 85.44, sales: 1009356, hours: 11811 },
    { month: "Sep 2024", productivity: 88.07, sales: 1015776, hours: 11535 },
    { month: "Oct 2024", productivity: 87.06, sales: 1048740, hours: 12046 },
    { month: "Nov 2024", productivity: 85.58, sales: 983604, hours: 11492 },
    { month: "Dec 2024", productivity: 84.89, sales: 936396, hours: 11031 },
    { month: "Jan 2025", productivity: 82.96, sales: 933221, hours: 11249 },
    { month: "Feb 2025", productivity: 92.44, sales: 950027, hours: 10278 },
    { month: "Mar 2025", productivity: 94.99, sales: 1040672, hours: 10954 },
    { month: "Apr 2025", productivity: 94.21, sales: 1033143, hours: 10967 },
    { month: "May 2025", productivity: 92.53, sales: 1060246, hours: 11459 },
    { month: "Jun 2025", productivity: 84.78, sales: 949035, hours: 11194 },
    { month: "Jul 2025", productivity: 84.21, sales: 900708, hours: 10697 },
    { month: "Aug 2025", productivity: 89.78, sales: 966739, hours: 10767 },
    { month: "Sep 2025", productivity: 93.47, sales: 929204, hours: 9941 }
  ];

  const [data, setData] = useState(initialData);
  const [newMonth, setNewMonth] = useState({ month: '', productivity: '', sales: '', hours: '' });
  const [showAddForm, setShowAddForm] = useState(false);
  const [editingIndex, setEditingIndex] = useState(null);
  const [editForm, setEditForm] = useState({ month: '', productivity: '', sales: '', hours: '' });
  const [deleteConfirm, setDeleteConfirm] = useState(null); // { index, month }
  const [showResetConfirm, setShowResetConfirm] = useState(false);
  const [isLoaded, setIsLoaded] = useState(false); // Track if data has been loaded from storage

  // Load data from persistent storage on mount
  React.useEffect(() => {
    const loadData = async () => {
      try {
        const stored = await window.storage.get('productivity-data');
        if (stored && stored.value) {
          setData(JSON.parse(stored.value));
        }
      } catch (error) {
        console.log('No stored data found, using initial data');
      }
      setIsLoaded(true); // Mark as loaded
    };
    loadData();
  }, []);

  // Save data to persistent storage whenever it changes (but only after initial load)
  React.useEffect(() => {
    if (!isLoaded) return; // Don't save on first render
    
    const saveData = async () => {
      try {
        await window.storage.set('productivity-data', JSON.stringify(data));
        console.log('Data saved successfully');
      } catch (error) {
        console.error('Error saving data:', error);
      }
    };
    saveData();
  }, [data, isLoaded]);

  const takeoverMonth = "Feb 2025";
  
  // Calculate statistics
  const getStats = () => {
    const takeoverIndex = data.findIndex(d => d.month === takeoverMonth);
    const beforeData = data.slice(0, takeoverIndex);
    const afterData = data.slice(takeoverIndex);
    
    const avg = arr => arr.reduce((sum, val) => sum + val, 0) / arr.length;
    
    const beforeAvg = avg(beforeData.map(d => d.productivity));
    const afterAvg = avg(afterData.map(d => d.productivity));
    const improvement = afterAvg - beforeAvg;
    const percentImprovement = (improvement / beforeAvg) * 100;
    
    const beforeGoalMet = beforeData.filter(d => d.productivity >= 90).length;
    const afterGoalMet = afterData.filter(d => d.productivity >= 90).length;
    
    return {
      beforeAvg: beforeAvg.toFixed(2),
      afterAvg: afterAvg.toFixed(2),
      improvement: improvement.toFixed(2),
      percentImprovement: percentImprovement.toFixed(1),
      beforeGoalPercent: ((beforeGoalMet / beforeData.length) * 100).toFixed(1),
      afterGoalPercent: ((afterGoalMet / afterData.length) * 100).toFixed(1),
      beforeCount: beforeData.length,
      afterCount: afterData.length,
      latestMonth: data[data.length - 1]
    };
  };

  const stats = getStats();

  const handleAddMonth = () => {
    if (newMonth.month && newMonth.productivity && newMonth.sales && newMonth.hours) {
      const newData = [...data, {
        month: newMonth.month,
        productivity: parseFloat(newMonth.productivity),
        sales: parseFloat(newMonth.sales),
        hours: parseFloat(newMonth.hours)
      }];
      setData(newData);
      setNewMonth({ month: '', productivity: '', sales: '', hours: '' });
      setShowAddForm(false);
    }
  };

  const handleDeleteMonth = (index) => {
    setDeleteConfirm({ index, month: data[index].month });
  };

  const confirmDelete = () => {
    if (deleteConfirm) {
      const newData = [...data];
      newData.splice(deleteConfirm.index, 1);
      setData(newData);
      setDeleteConfirm(null);
    }
  };

  const cancelDelete = () => {
    setDeleteConfirm(null);
  };

  const handleEditMonth = (index) => {
    setEditingIndex(index);
    setEditForm({
      month: data[index].month,
      productivity: data[index].productivity.toString(),
      sales: data[index].sales.toString(),
      hours: data[index].hours.toString()
    });
  };

  const handleSaveEdit = () => {
    if (editForm.month && editForm.productivity && editForm.sales && editForm.hours) {
      const newData = [...data];
      newData[editingIndex] = {
        month: editForm.month,
        productivity: parseFloat(editForm.productivity),
        sales: parseFloat(editForm.sales),
        hours: parseFloat(editForm.hours)
      };
      setData(newData);
      setEditingIndex(null);
      setEditForm({ month: '', productivity: '', sales: '', hours: '' });
    }
  };

  const handleCancelEdit = () => {
    setEditingIndex(null);
    setEditForm({ month: '', productivity: '', sales: '', hours: '' });
  };

  const resetToOriginal = async () => {
    setShowResetConfirm(true);
  };

  const confirmReset = async () => {
    setData(initialData);
    try {
      await window.storage.set('productivity-data', JSON.stringify(initialData));
    } catch (error) {
      console.error('Error resetting data:', error);
    }
    setShowResetConfirm(false);
  };

  const cancelReset = () => {
    setShowResetConfirm(false);
  };

  const exportData = () => {
    // Create CSV content
    const headers = ['Month,Productivity,Total Sales,Total Hours'];
    const rows = data.map(row => 
      `${row.month},${row.productivity},${row.sales},${row.hours}`
    );
    const csvContent = headers.concat(rows).join('\n');
    
    const dataBlob = new Blob([csvContent], { type: 'text/csv' });
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `productivity-data-${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const importData = (event) => {
    const file = event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const csvContent = e.target.result;
          const parsed = Papa.parse(csvContent, {
            header: true,
            dynamicTyping: true,
            skipEmptyLines: true
          });
          
          // Convert parsed CSV to our data format
          const imported = parsed.data.map(row => ({
            month: row.Month || row.month,
            productivity: parseFloat(row.Productivity || row.productivity),
            sales: parseFloat(row['Total Sales'] || row.sales),
            hours: parseFloat(row['Total Hours'] || row.hours)
          }));
          
          setData(imported);
          alert('Data imported successfully!');
        } catch (error) {
          alert('Error importing CSV. Please check the file format.');
          console.error('Import error:', error);
        }
      };
      reader.readAsText(file);
    }
  };

  return (
    <div className="w-full max-w-7xl mx-auto p-6 bg-gray-50 min-h-screen">
      <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-3xl font-bold text-gray-800 flex items-center gap-2">
              <TrendingUp className="text-green-600" />
              Productivity Performance Tracker
            </h1>
            <p className="text-gray-600 mt-1">Supervision began: {takeoverMonth}</p>
          </div>
          <div className="flex gap-2">
            <button
              onClick={() => setShowAddForm(!showAddForm)}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition"
            >
              <Plus size={20} />
              Add Month
            </button>
            <button
              onClick={exportData}
              className="flex items-center gap-2 px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition"
            >
              <Download size={20} />
              Export
            </button>
            <label className="flex items-center gap-2 px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 transition cursor-pointer">
              <Upload size={20} />
              Import
              <input type="file" accept=".csv" onChange={importData} className="hidden" />
            </label>
            <button
              onClick={resetToOriginal}
              className="flex items-center gap-2 px-4 py-2 bg-orange-600 text-white rounded-lg hover:bg-orange-700 transition"
            >
              Reset
            </button>
          </div>
        </div>

        {showAddForm && (
          <div className="bg-blue-50 p-4 rounded-lg mb-6 border border-blue-200">
            <h3 className="font-semibold mb-3 text-gray-800">Add New Month</h3>
            <div className="grid grid-cols-4 gap-3">
              <input
                type="text"
                placeholder="Month (e.g., Oct 2025)"
                value={newMonth.month}
                onChange={(e) => setNewMonth({ ...newMonth, month: e.target.value })}
                className="px-3 py-2 border border-gray-300 rounded"
              />
              <input
                type="number"
                step="0.01"
                placeholder="Productivity"
                value={newMonth.productivity}
                onChange={(e) => setNewMonth({ ...newMonth, productivity: e.target.value })}
                className="px-3 py-2 border border-gray-300 rounded"
              />
              <input
                type="number"
                step="0.01"
                placeholder="Total Sales"
                value={newMonth.sales}
                onChange={(e) => setNewMonth({ ...newMonth, sales: e.target.value })}
                className="px-3 py-2 border border-gray-300 rounded"
              />
              <input
                type="number"
                step="0.01"
                placeholder="Total Hours"
                value={newMonth.hours}
                onChange={(e) => setNewMonth({ ...newMonth, hours: e.target.value })}
                className="px-3 py-2 border border-gray-300 rounded"
              />
            </div>
            <button
              onClick={handleAddMonth}
              className="mt-3 px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 transition"
            >
              Add to Database
            </button>
          </div>
        )}

        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-gradient-to-br from-blue-500 to-blue-600 text-white p-4 rounded-lg shadow">
            <div className="text-sm opacity-90">Before Supervision</div>
            <div className="text-3xl font-bold">{stats.beforeAvg}</div>
            <div className="text-xs opacity-75 mt-1">{stats.beforeCount} months</div>
          </div>
          
          <div className="bg-gradient-to-br from-green-500 to-green-600 text-white p-4 rounded-lg shadow">
            <div className="text-sm opacity-90">Under Your Supervision</div>
            <div className="text-3xl font-bold">{stats.afterAvg}</div>
            <div className="text-xs opacity-75 mt-1">{stats.afterCount} months</div>
          </div>
          
          <div className="bg-gradient-to-br from-purple-500 to-purple-600 text-white p-4 rounded-lg shadow">
            <div className="text-sm opacity-90">Improvement</div>
            <div className="text-3xl font-bold">+{stats.improvement}</div>
            <div className="text-xs opacity-75 mt-1">{stats.percentImprovement}% increase</div>
          </div>
          
          <div className="bg-gradient-to-br from-orange-500 to-orange-600 text-white p-4 rounded-lg shadow">
            <div className="text-sm opacity-90">Latest Month</div>
            <div className="text-3xl font-bold">{stats.latestMonth.productivity.toFixed(2)}</div>
            <div className="text-xs opacity-75 mt-1">{stats.latestMonth.month}</div>
          </div>
        </div>

        <div className="bg-gray-50 p-4 rounded-lg mb-6 border border-gray-200">
          <h3 className="font-semibold mb-2 text-gray-800">Goal Achievement (90+ Target)</h3>
          <div className="grid grid-cols-2 gap-4">
            <div>
              <div className="text-sm text-gray-600">Before: {stats.beforeGoalPercent}% of months</div>
              <div className="w-full bg-gray-200 rounded-full h-3 mt-1">
                <div
                  className="bg-blue-500 h-3 rounded-full transition-all"
                  style={{ width: `${stats.beforeGoalPercent}%` }}
                />
              </div>
            </div>
            <div>
              <div className="text-sm text-gray-600">After: {stats.afterGoalPercent}% of months</div>
              <div className="w-full bg-gray-200 rounded-full h-3 mt-1">
                <div
                  className="bg-green-500 h-3 rounded-full transition-all"
                  style={{ width: `${stats.afterGoalPercent}%` }}
                />
              </div>
            </div>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
        <h2 className="text-xl font-bold mb-4 text-gray-800">Productivity Trend</h2>
        <ResponsiveContainer width="100%" height={400}>
          <LineChart data={data}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis
              dataKey="month"
              angle={-45}
              textAnchor="end"
              height={100}
              interval={Math.floor(data.length / 20)}
            />
            <YAxis domain={[60, 100]} />
            <Tooltip />
            <Legend />
            <ReferenceLine y={90} stroke="#ef4444" strokeDasharray="3 3" label="Goal: 90" />
            <ReferenceLine
              x={takeoverMonth}
              stroke="#10b981"
              strokeWidth={2}
              label={{ value: 'Supervision Started', position: 'top' }}
            />
            <Line
              type="monotone"
              dataKey="productivity"
              stroke="#3b82f6"
              strokeWidth={2}
              dot={{ fill: '#3b82f6', r: 3 }}
              activeDot={{ r: 6 }}
            />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <div className="bg-white rounded-lg shadow-lg p-6">
        <h2 className="text-xl font-bold mb-4 text-gray-800">Monthly Details</h2>
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-gray-100">
              <tr>
                <th className="px-4 py-2 text-left">Month</th>
                <th className="px-4 py-2 text-right">Productivity</th>
                <th className="px-4 py-2 text-right">Total Sales</th>
                <th className="px-4 py-2 text-right">Total Hours</th>
                <th className="px-4 py-2 text-center">Goal Met</th>
                <th className="px-4 py-2 text-center">Actions</th>
              </tr>
            </thead>
            <tbody>
              {data.slice().reverse().map((row, idx) => {
                const actualIndex = data.length - 1 - idx;
                const isEditing = editingIndex === actualIndex;
                
                return (
                  <tr
                    key={actualIndex}
                    className={`border-b ${row.month === takeoverMonth ? 'bg-green-50 border-green-300' : ''}`}
                  >
                    {isEditing ? (
                      <>
                        <td className="px-4 py-2">
                          <input
                            type="text"
                            value={editForm.month}
                            onChange={(e) => setEditForm({ ...editForm, month: e.target.value })}
                            className="w-full px-2 py-1 border border-gray-300 rounded"
                          />
                        </td>
                        <td className="px-4 py-2">
                          <input
                            type="number"
                            step="0.01"
                            value={editForm.productivity}
                            onChange={(e) => setEditForm({ ...editForm, productivity: e.target.value })}
                            className="w-full px-2 py-1 border border-gray-300 rounded text-right"
                          />
                        </td>
                        <td className="px-4 py-2">
                          <input
                            type="number"
                            value={editForm.sales}
                            onChange={(e) => setEditForm({ ...editForm, sales: e.target.value })}
                            className="w-full px-2 py-1 border border-gray-300 rounded text-right"
                          />
                        </td>
                        <td className="px-4 py-2">
                          <input
                            type="number"
                            step="0.01"
                            value={editForm.hours}
                            onChange={(e) => setEditForm({ ...editForm, hours: e.target.value })}
                            className="w-full px-2 py-1 border border-gray-300 rounded text-right"
                          />
                        </td>
                        <td className="px-4 py-2 text-center">
                          {parseFloat(editForm.productivity) >= 90 ? (
                            <span className="text-green-600 font-bold">✓</span>
                          ) : (
                            <span className="text-gray-400">—</span>
                          )}
                        </td>
                        <td className="px-4 py-2 text-center">
                          <div className="flex gap-2 justify-center">
                            <button
                              onClick={handleSaveEdit}
                              className="text-green-600 hover:text-green-800 hover:bg-green-50 p-1 rounded transition"
                              title="Save changes"
                            >
                              <Check size={18} />
                            </button>
                            <button
                              onClick={handleCancelEdit}
                              className="text-gray-600 hover:text-gray-800 hover:bg-gray-50 p-1 rounded transition"
                              title="Cancel"
                            >
                              <X size={18} />
                            </button>
                          </div>
                        </td>
                      </>
                    ) : (
                      <>
                        <td className="px-4 py-2 font-medium">{row.month}</td>
                        <td className="px-4 py-2 text-right">{row.productivity.toFixed(2)}</td>
                        <td className="px-4 py-2 text-right">${row.sales.toLocaleString()}</td>
                        <td className="px-4 py-2 text-right">{row.hours.toLocaleString()}</td>
                        <td className="px-4 py-2 text-center">
                          {row.productivity >= 90 ? (
                            <span className="text-green-600 font-bold">✓</span>
                          ) : (
                            <span className="text-gray-400">—</span>
                          )}
                        </td>
                        <td className="px-4 py-2 text-center">
                          <div className="flex gap-2 justify-center">
                            <button
                              onClick={() => handleEditMonth(actualIndex)}
                              className="text-blue-600 hover:text-blue-800 hover:bg-blue-50 p-1 rounded transition"
                              title="Edit this month"
                            >
                              <Edit2 size={18} />
                            </button>
                            <button
                              onClick={() => handleDeleteMonth(actualIndex)}
                              className="text-red-600 hover:text-red-800 hover:bg-red-50 p-1 rounded transition"
                              title="Delete this month"
                            >
                              <Trash2 size={18} />
                            </button>
                          </div>
                        </td>
                      </>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Delete Confirmation Modal */}
      {deleteConfirm && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg shadow-xl p-6 max-w-md">
            <h3 className="text-xl font-bold mb-4">Confirm Delete</h3>
            <p className="mb-6">Are you sure you want to delete <strong>{deleteConfirm.month}</strong>?</p>
            <div className="flex gap-3 justify-end">
              <button
                onClick={cancelDelete}
                className="px-4 py-2 bg-gray-200 text-gray-800 rounded hover:bg-gray-300 transition"
              >
                Cancel
              </button>
              <button
                onClick={confirmDelete}
                className="px-4 py-2 bg-red-600 text-white rounded hover:bg-red-700 transition"
              >
                Delete
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Reset Confirmation Modal */}
      {showResetConfirm && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg shadow-xl p-6 max-w-md">
            <h3 className="text-xl font-bold mb-4">Confirm Reset</h3>
            <p className="mb-6">Reset to original data? This will delete all added months.</p>
            <div className="flex gap-3 justify-end">
              <button
                onClick={cancelReset}
                className="px-4 py-2 bg-gray-200 text-gray-800 rounded hover:bg-gray-300 transition"
              >
                Cancel
              </button>
              <button
                onClick={confirmReset}
                className="px-4 py-2 bg-orange-600 text-white rounded hover:bg-orange-700 transition"
              >
                Reset
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ProductivityTracker;