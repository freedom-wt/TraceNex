/*
Copyright (C) 2025 QuantumNous

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as
published by the Free Software Foundation, either version 3 of the
License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program. If not, see <https://www.gnu.org/licenses/>.

For commercial licensing, please contact support@quantumnous.com
*/

import { useState, useEffect, useCallback } from 'react';
import { useTranslation } from 'react-i18next';
import { Modal } from '@douyinfe/semi-ui';
import {
  API,
  getTodayStartTimestamp,
  isAdmin,
  showError,
  showSuccess,
  timestamp2string,
  renderQuota,
  renderNumber,
  getLogOther,
  copy,
  renderClaudeLogContent,
  renderLogContent,
  renderAudioModelPrice,
  renderClaudeModelPrice,
  renderModelPrice,
} from '../../helpers';
import { ITEMS_PER_PAGE } from '../../constants';
import { useTableCompactMode } from '../common/useTableCompactMode';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export const useLogsData = () => {
  const { t } = useTranslation();

  // Define column keys for selection
  const COLUMN_KEYS = {
    TIME: 'time',
    CHANNEL: 'channel',
    USERNAME: 'username',
    TOKEN: 'token',
    GROUP: 'group',
    TYPE: 'type',
    MODEL: 'model',
    USE_TIME: 'use_time',
    PROMPT: 'prompt',
    COMPLETION: 'completion',
    COST: 'cost',
    RETRY: 'retry',
    IP: 'ip',
    DETAILS: 'details',
  };

  // Basic state
  const [logs, setLogs] = useState([]);
  const [expandData, setExpandData] = useState({});
  const [showStat, setShowStat] = useState(false);
  const [loading, setLoading] = useState(false);
  const [loadingStat, setLoadingStat] = useState(false);
  const [activePage, setActivePage] = useState(1);
  const [logCount, setLogCount] = useState(0);
  const [pageSize, setPageSize] = useState(ITEMS_PER_PAGE);
  const [logType, setLogType] = useState(0);

  // User and admin
  const isAdminUser = isAdmin();
  // Role-specific storage key to prevent different roles from overwriting each other
  const STORAGE_KEY = isAdminUser
    ? 'logs-table-columns-admin'
    : 'logs-table-columns-user';

  // Statistics state
  const [stat, setStat] = useState({
    quota: 0,
    token: 0,
  });

  // Form state
  const [formApi, setFormApi] = useState(null);
  let now = new Date();
  const formInitValues = {
    username: '',
    token_name: '',
    model_name: '',
    channel: '',
    group: '',
    dateRange: [
      timestamp2string(getTodayStartTimestamp()),
      timestamp2string(now.getTime() / 1000 + 3600),
    ],
    logType: '0',
  };

  // Column visibility state
  const [visibleColumns, setVisibleColumns] = useState({});
  const [showColumnSelector, setShowColumnSelector] = useState(false);

  // Compact mode
  const [compactMode, setCompactMode] = useTableCompactMode('logs');

  // User info modal state
  const [showUserInfo, setShowUserInfoModal] = useState(false);
  const [userInfoData, setUserInfoData] = useState(null);

  // Load saved column preferences from localStorage
  useEffect(() => {
    const savedColumns = localStorage.getItem(STORAGE_KEY);
    if (savedColumns) {
      try {
        const parsed = JSON.parse(savedColumns);
        const defaults = getDefaultColumnVisibility();
        const merged = { ...defaults, ...parsed };

        // For non-admin users, force-hide admin-only columns (does not touch admin settings)
        if (!isAdminUser) {
          merged[COLUMN_KEYS.CHANNEL] = false;
          merged[COLUMN_KEYS.USERNAME] = false;
          merged[COLUMN_KEYS.RETRY] = false;
        }
        setVisibleColumns(merged);
      } catch (e) {
        console.error('Failed to parse saved column preferences', e);
        initDefaultColumns();
      }
    } else {
      initDefaultColumns();
    }
  }, []);

  // Get default column visibility based on user role
  const getDefaultColumnVisibility = () => {
    return {
      [COLUMN_KEYS.TIME]: true,
      [COLUMN_KEYS.CHANNEL]: isAdminUser,
      [COLUMN_KEYS.USERNAME]: isAdminUser,
      [COLUMN_KEYS.TOKEN]: true,
      [COLUMN_KEYS.GROUP]: true,
      [COLUMN_KEYS.TYPE]: true,
      [COLUMN_KEYS.MODEL]: true,
      [COLUMN_KEYS.USE_TIME]: true,
      [COLUMN_KEYS.PROMPT]: true,
      [COLUMN_KEYS.COMPLETION]: true,
      [COLUMN_KEYS.COST]: true,
      [COLUMN_KEYS.RETRY]: isAdminUser,
      [COLUMN_KEYS.IP]: true,
      [COLUMN_KEYS.DETAILS]: true,
    };
  };

  // Initialize default column visibility
  const initDefaultColumns = () => {
    const defaults = getDefaultColumnVisibility();
    setVisibleColumns(defaults);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(defaults));
  };

  // Handle column visibility change
  const handleColumnVisibilityChange = (columnKey, checked) => {
    const updatedColumns = { ...visibleColumns, [columnKey]: checked };
    setVisibleColumns(updatedColumns);
  };

  // Handle "Select All" checkbox
  const handleSelectAll = (checked) => {
    const allKeys = Object.keys(COLUMN_KEYS).map((key) => COLUMN_KEYS[key]);
    const updatedColumns = {};

    allKeys.forEach((key) => {
      if (
        (key === COLUMN_KEYS.CHANNEL ||
          key === COLUMN_KEYS.USERNAME ||
          key === COLUMN_KEYS.RETRY) &&
        !isAdminUser
      ) {
        updatedColumns[key] = false;
      } else {
        updatedColumns[key] = checked;
      }
    });

    setVisibleColumns(updatedColumns);
  };

  // Persist column settings to the role-specific STORAGE_KEY
  useEffect(() => {
    if (Object.keys(visibleColumns).length > 0) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(visibleColumns));
    }
  }, [visibleColumns]);

  // 获取表单值的辅助函数，确保所有值都是字符串
  const getFormValues = () => {
    const formValues = formApi ? formApi.getValues() : {};

    let start_timestamp = timestamp2string(getTodayStartTimestamp());
    let end_timestamp = timestamp2string(now.getTime() / 1000 + 3600);

    if (
      formValues.dateRange &&
      Array.isArray(formValues.dateRange) &&
      formValues.dateRange.length === 2
    ) {
      start_timestamp = formValues.dateRange[0];
      end_timestamp = formValues.dateRange[1];
    }

    return {
      username: formValues.username || '',
      token_name: formValues.token_name || '',
      model_name: formValues.model_name || '',
      start_timestamp,
      end_timestamp,
      channel: formValues.channel || '',
      group: formValues.group || '',
      logType: formValues.logType ? parseInt(formValues.logType) : 0,
    };
  };

  // Statistics functions
  const getLogSelfStat = async () => {
    const {
      token_name,
      model_name,
      start_timestamp,
      end_timestamp,
      group,
      logType: formLogType,
    } = getFormValues();
    const currentLogType = formLogType !== undefined ? formLogType : logType;
    let localStartTimestamp = Date.parse(start_timestamp) / 1000;
    let localEndTimestamp = Date.parse(end_timestamp) / 1000;
    let url = `/api/log/self/stat?type=${currentLogType}&token_name=${token_name}&model_name=${model_name}&start_timestamp=${localStartTimestamp}&end_timestamp=${localEndTimestamp}&group=${group}`;
    url = encodeURI(url);
    let res = await API.get(url);
    const { success, message, data } = res.data;
    if (success) {
      setStat(data);
    } else {
      showError(message);
    }
  };

  const getLogStat = async () => {
    const {
      username,
      token_name,
      model_name,
      start_timestamp,
      end_timestamp,
      channel,
      group,
      logType: formLogType,
    } = getFormValues();
    const currentLogType = formLogType !== undefined ? formLogType : logType;
    let localStartTimestamp = Date.parse(start_timestamp) / 1000;
    let localEndTimestamp = Date.parse(end_timestamp) / 1000;
    let url = `/api/log/stat?type=${currentLogType}&username=${username}&token_name=${token_name}&model_name=${model_name}&start_timestamp=${localStartTimestamp}&end_timestamp=${localEndTimestamp}&channel=${channel}&group=${group}`;
    url = encodeURI(url);
    let res = await API.get(url);
    const { success, message, data } = res.data;
    if (success) {
      setStat(data);
    } else {
      showError(message);
    }
  };

  const handleEyeClick = async () => {
    if (loadingStat) {
      return;
    }
    setLoadingStat(true);
    if (isAdminUser) {
      await getLogStat();
    } else {
      await getLogSelfStat();
    }
    setShowStat(true);
    setLoadingStat(false);
  };

  // User info function
  const showUserInfoFunc = async (userId) => {
    if (!isAdminUser) {
      return;
    }
    const res = await API.get(`/api/user/${userId}`);
    const { success, message, data } = res.data;
    if (success) {
      setUserInfoData(data);
      setShowUserInfoModal(true);
    } else {
      showError(message);
    }
  };

  // Format logs data
  const setLogsFormat = (logs) => {
    let expandDatesLocal = {};
    for (let i = 0; i < logs.length; i++) {
      logs[i].timestamp2string = timestamp2string(logs[i].created_at);
      logs[i].key = logs[i].id;
      let other = getLogOther(logs[i].other);
      let expandDataLocal = [];

      if (isAdminUser && (logs[i].type === 0 || logs[i].type === 2)) {
        expandDataLocal.push({
          key: t('渠道信息'),
          value: `${logs[i].channel} - ${logs[i].channel_name || '[未知]'}`,
        });
      }
      if (other?.ws || other?.audio) {
        expandDataLocal.push({
          key: t('语音输入'),
          value: other.audio_input,
        });
        expandDataLocal.push({
          key: t('语音输出'),
          value: other.audio_output,
        });
        expandDataLocal.push({
          key: t('文字输入'),
          value: other.text_input,
        });
        expandDataLocal.push({
          key: t('文字输出'),
          value: other.text_output,
        });
      }
      if (other?.cache_tokens > 0) {
        expandDataLocal.push({
          key: t('缓存 Tokens'),
          value: other.cache_tokens,
        });
      }
      if (other?.cache_creation_tokens > 0) {
        expandDataLocal.push({
          key: t('缓存创建 Tokens'),
          value: other.cache_creation_tokens,
        });
      }
      if (logs[i].type === 2) {
        expandDataLocal.push({
          key: t('日志详情'),
          value: other?.claude
            ? renderClaudeLogContent(
              other?.model_ratio,
              other.completion_ratio,
              other.model_price,
              other.group_ratio,
              other?.user_group_ratio,
              other.cache_ratio || 1.0,
              other.cache_creation_ratio || 1.0,
              other.cache_creation_tokens_5m || 0,
              other.cache_creation_ratio_5m || other.cache_creation_ratio || 1.0,
              other.cache_creation_tokens_1h || 0,
              other.cache_creation_ratio_1h || other.cache_creation_ratio || 1.0,
            )
            : renderLogContent(
              other?.model_ratio,
              other.completion_ratio,
              other.model_price,
              other.group_ratio,
              other?.user_group_ratio,
              other.cache_ratio || 1.0,
              false,
              1.0,
              other.web_search || false,
              other.web_search_call_count || 0,
              other.file_search || false,
              other.file_search_call_count || 0,
            ),
        });
        if (logs[i]?.content) {
          expandDataLocal.push({
            key: t('其他详情'),
            value: logs[i].content,
          });
        }
      }
      if (logs[i].type === 2) {
        let modelMapped =
          other?.is_model_mapped &&
          other?.upstream_model_name &&
          other?.upstream_model_name !== '';
        if (modelMapped) {
          expandDataLocal.push({
            key: t('请求并计费模型'),
            value: logs[i].model_name,
          });
          expandDataLocal.push({
            key: t('实际模型'),
            value: other.upstream_model_name,
          });
        }
        let content = '';
        if (other?.ws || other?.audio) {
          content = renderAudioModelPrice(
            other?.text_input,
            other?.text_output,
            other?.model_ratio,
            other?.model_price,
            other?.completion_ratio,
            other?.audio_input,
            other?.audio_output,
            other?.audio_ratio,
            other?.audio_completion_ratio,
            other?.group_ratio,
            other?.user_group_ratio,
            other?.cache_tokens || 0,
            other?.cache_ratio || 1.0,
          );
        } else if (other?.claude) {
          content = renderClaudeModelPrice(
            logs[i].prompt_tokens,
            logs[i].completion_tokens,
            other.model_ratio,
            other.model_price,
            other.completion_ratio,
            other.group_ratio,
            other?.user_group_ratio,
            other.cache_tokens || 0,
            other.cache_ratio || 1.0,
            other.cache_creation_tokens || 0,
            other.cache_creation_ratio || 1.0,
            other.cache_creation_tokens_5m || 0,
            other.cache_creation_ratio_5m || other.cache_creation_ratio || 1.0,
            other.cache_creation_tokens_1h || 0,
            other.cache_creation_ratio_1h || other.cache_creation_ratio || 1.0,
          );
        } else {
          content = renderModelPrice(
            logs[i].prompt_tokens,
            logs[i].completion_tokens,
            other?.model_ratio,
            other?.model_price,
            other?.completion_ratio,
            other?.group_ratio,
            other?.user_group_ratio,
            other?.cache_tokens || 0,
            other?.cache_ratio || 1.0,
            other?.image || false,
            other?.image_ratio || 0,
            other?.image_output || 0,
            other?.web_search || false,
            other?.web_search_call_count || 0,
            other?.web_search_price || 0,
            other?.file_search || false,
            other?.file_search_call_count || 0,
            other?.file_search_price || 0,
            other?.audio_input_seperate_price || false,
            other?.audio_input_token_count || 0,
            other?.audio_input_price || 0,
            other?.image_generation_call || false,
            other?.image_generation_call_price || 0,
          );
        }
        expandDataLocal.push({
          key: t('计费过程'),
          value: content,
        });
        if (other?.reasoning_effort) {
          expandDataLocal.push({
            key: t('Reasoning Effort'),
            value: other.reasoning_effort,
          });
        }
      }
      if (other?.request_path) {
        expandDataLocal.push({
          key: t('请求路径'),
          value: other.request_path,
        });
      }
      if (isAdminUser) {
        let localCountMode = '';
        if (other?.admin_info?.local_count_tokens) {
          localCountMode = t('本地计费');
        } else {
          localCountMode = t('上游返回');
        }
        expandDataLocal.push({
          key: t('计费模式'),
          value: localCountMode,
        });
      }
      expandDatesLocal[logs[i].key] = expandDataLocal;
    }

    setExpandData(expandDatesLocal);
    setLogs(logs);
  };

  const [exportProgress,setExportProgress] = useState('');
  const [isExporting, setIsExporting] = useState(false);
  const [abortController, setAbortController] = useState(null);

  const fetchPageData = async (pageNum,customLogType = null) => {
    let url = '';
    const {
      username,
      token_name,
      model_name,
      start_timestamp,
      end_timestamp,
      channel,
      group,
      logType: formLogType,
    } = getFormValues();
    
    const currentLogType =
      customLogType !== null
        ? customLogType
        : formLogType !== undefined
          ? formLogType
          : logType;

    let localStartTimestamp = Date.parse(start_timestamp) / 1000;
    let localEndTimestamp = Date.parse(end_timestamp) / 1000;

    if (isAdminUser) {
      url = `/api/log/?p=${pageNum}&page_size=${100}&type=${currentLogType}&username=${username}&token_name=${token_name}&model_name=${model_name}&start_timestamp=${localStartTimestamp}&end_timestamp=${localEndTimestamp}&channel=${channel}&group=${group}`;
    } else {
      url = `/api/log/self/?p=${pageNum}&page_size=${100}&type=${currentLogType}&token_name=${token_name}&model_name=${model_name}&start_timestamp=${localStartTimestamp}&end_timestamp=${localEndTimestamp}&group=${group}`;
    }
    
    url = encodeURI(url);
    const res = await API.get(url);
    const { success, message, data } = res.data;
    if (success) {
      const newPageData = data.items;
      setActivePage(data.page);
      setPageSize(data.page_size);
      setLogCount(data.total);

      return {
        list: newPageData,
        total: data.total,
      }
    } else {
      showError(message);
    }
  }

 // 辅助函数1：时间戳转 YYYY-MM-DD HH:mm:ss 格式
  const formatTimestamp = (timestamp) => {
    if (!timestamp) return '无';
    const date = new Date(timestamp * 1000);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
  };

  // 辅助函数2：类型数字转文字描述（适配你的接口 type 字段）
  const formatLogType = (type) => {
    const typeMap = {
      2: '消费',
      3: '额度修改',
      // 可补充其他 type 对应的文字
    };
    return typeMap[type] || `未知类型(${type})`;
  };

  // 辅助函数3：解析 other 字段，提取关键信息
  const parseOtherField = (otherStr) => {
    if (!otherStr) return { model_ratio: 0, group_ratio: 0, admin_info: {} };
    try {
      const otherObj = JSON.parse(otherStr);
      return {
        model_ratio: otherObj.model_ratio || 0,
        group_ratio: otherObj.group_ratio || 0,
        admin_info: otherObj.admin_info || {},
      };
    } catch (e) {
      return { model_ratio: 0, group_ratio: 0, admin_info: {} };
    }
  };

  // 辅助函数4：计算花费（基于 token 数 + 模型倍率，可适配你的业务规则）
  const calculateCost = (promptTokens, completionTokens, modelRatio, groupRatio) => {
    // 基础费率（可根据你的业务调整，单位：元/1000 token）
    const basePromptRate = 0.01; // 输入 token 基础费率
    const baseCompletionRate = 0.03; // 输出 token 基础费率

    if (promptTokens === 0 && completionTokens === 0) return '0.000000';

    // 实际花费 = (输入token×输入基础费率 + 输出token×输出基础费率) × 模型倍率 × 分组倍率
    const promptCost = (promptTokens / 1000) * basePromptRate * modelRatio * groupRatio;
    const completionCost = (completionTokens / 1000) * baseCompletionRate * modelRatio * groupRatio;
    return (promptCost + completionCost).toFixed(6);
  };

  // 辅助函数5：处理单条原始数据，映射为 Excel 所需格式
  const formatSingleRow = (rawItem) => {
    const { model_ratio, group_ratio } = parseOtherField(rawItem.other);
    const isStreamText = rawItem.is_stream ? '流' : '非流';

    return {
      time: formatTimestamp(rawItem.created_at),
      channel: rawItem.channel_name || rawItem.channel || '无',
      token: rawItem.token_name || '无',
      group: rawItem.group || '无',
      type: formatLogType(rawItem.type),
      model: rawItem.model_name || '无',
      use_time: rawItem.use_time > 0 ? `${rawItem.use_time} s ${isStreamText}` : `0 s ${isStreamText}`,
      prompt: rawItem.prompt_tokens || 0,
      completion: rawItem.completion_tokens || 0,
      cost: calculateCost(rawItem.prompt_tokens, rawItem.completion_tokens, model_ratio, group_ratio),
      ip: rawItem.ip || '无',
      retry: '无', // 你的接口返回中无 retry 字段，先兜底为「无」，后续可补充
      details: `模型:${model_ratio} * 分组倍率:${group_ratio}` // 匹配图一详情格式
    };
  };

  const exportAllLogs = useCallback(async () => {
    
    if (isExporting) return;
    setIsExporting(true);
    setExportProgress('开始初始化导出...');

    const controller = new AbortController();
    setAbortController(controller);
    const { signal } = controller;

    try {
      // 配置项
      const pageSize = 100; // 每页拉取条数
      let currentPage = 1; // 当前请求页码
      let totalPages = 1; // 总页数（初始值，后续计算）
      let totalCount = 0; //总条数（从第一页接口获取）
      let fetchedTotal = 0; // 已拉取的总条数

      //初始化Excel
      const workBook = new ExcelJS.Workbook();
      const workSheet = workBook.addWorksheet('全量数据',{
        views:[{state: 'frozen', ySplit: 1 }], // 冻结表头
      });

      // 设置Excel表头（适配数据字段）
      workSheet.columns = [
        {header: '时间', key: 'time', width: 20},
        {header: '渠道', key: 'channel', width: 10},
        // {header: '用户', key: 'username', width: 10},
        {header: '令牌', key: 'token', width: 10},
        {header: '分组', key: 'group', width: 10},
        {header: '类型', key: 'type', width: 10},
        {header: '模型', key: 'model', width: 20},
        {header: '用时/首字', key: 'use_time', width: 10},
        {header: '输入', key: 'prompt', width: 10},
        {header: '输出', key: 'completion', width: 10},
        {header: '花费', key: 'cost', width: 10},
        {header: 'IP', key: 'ip', width: 10},
        {header: '重试', key: 'retry', width: 10},
        {header: '详情', key: 'details', width: 30},
      ];
      // 表头样式
      workSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      workSheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3366FF' } };

      // 循环拉取所有分页数据
      while (true) {
        // 检查是否中断
        if (signal.aborted) break;
 
        // 拉取当前页数据
        setExportProgress(`正在获取第${currentPage}页数据...`);
        const pageResult = await fetchPageData(currentPage);

        if (!pageResult || !pageResult.list || pageResult.list.length === 0) break;
        // 第一页数据：获取总条数并计算总页数
        if (currentPage === 1) {
          totalCount = pageResult.total;
          // 计算总页数：向上取整
          totalPages = Math.ceil(totalCount / pageSize);
          setExportProgress(`共${totalCount}条数据，分${totalPages}页，正在获取第${currentPage}页...`);
        }
        
        // 关键修改：处理当前页原始数据，映射为Excel所需格式
        const formattedRows = pageResult.list.map(item => formatSingleRow(item));

        // 追加当前页数据到Excel
        workSheet.addRows(formattedRows);
        fetchedTotal += pageResult.list.length;

        // 更新进度
        const progress = Math.floor((currentPage / totalPages) * 100);
        setExportProgress(`已获取${currentPage}/${totalPages}页（共${fetchedTotal}/${totalCount}条），进度${progress}%`);

        // 让出主线程
        await new Promise(resolve => setTimeout(resolve, 0));

        // 页码自增，判断是否到最后一页
        currentPage++;
        if (currentPage > totalPages) break;
      }

      // 检查是否被主动中断
      if (signal.aborted) {
        setExportProgress('导出已取消');
        return;
      }

      //生成并下载Excel
      setExportProgress('正在生成Excel文件...');
      const buffer = await workBook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      saveAs(blob, `全量数据_${new Date().toLocaleDateString()}.xlsx`);
      setExportProgress(`导出完成！共导出${fetchedTotal}条数据（总数据量${totalCount}条）`)
    } catch (error) {
      if (error.name !== 'AbortError') {
        setExportProgress(`导出失败：${error.message}`);
        alert(`导出失败：${error.message}`);
      }
    } finally {
      setTimeout(() => {
        setIsExporting(false);
        setAbortController(null);
        if (!signal.aborted) {
          setTimeout(() => setExportProgress(''), 3000)
        }
      }, 1000)
    }
  }, [isExporting]);

  // 取消导出
  const cancelExport = () => {
    if (abortController) {
      abortController.abort();
      setExportProgress('正在取消导出...');
    }
  };

  // Load logs function
  const loadLogs = async (startIdx, pageSize, customLogType = null) => {
    setLoading(true);

    let url = '';
    const {
      username,
      token_name,
      model_name,
      start_timestamp,
      end_timestamp,
      channel,
      group,
      logType: formLogType,
    } = getFormValues();

    const currentLogType =
      customLogType !== null
        ? customLogType
        : formLogType !== undefined
          ? formLogType
          : logType;

    let localStartTimestamp = Date.parse(start_timestamp) / 1000;
    let localEndTimestamp = Date.parse(end_timestamp) / 1000;
    if (isAdminUser) {
      url = `/api/log/?p=${startIdx}&page_size=${pageSize}&type=${currentLogType}&username=${username}&token_name=${token_name}&model_name=${model_name}&start_timestamp=${localStartTimestamp}&end_timestamp=${localEndTimestamp}&channel=${channel}&group=${group}`;
    } else {
      url = `/api/log/self/?p=${startIdx}&page_size=${pageSize}&type=${currentLogType}&token_name=${token_name}&model_name=${model_name}&start_timestamp=${localStartTimestamp}&end_timestamp=${localEndTimestamp}&group=${group}`;
    }
    url = encodeURI(url);
    const res = await API.get(url);
    const { success, message, data } = res.data;
    if (success) {
      const newPageData = data.items;
      setActivePage(data.page);
      setPageSize(data.page_size);
      setLogCount(data.total);

      setLogsFormat(newPageData);
    } else {
      showError(message);
    }
    setLoading(false);
  };

  // Page handlers
  const handlePageChange = (page) => {
    setActivePage(page);
    loadLogs(page, pageSize).then((r) => { });
  };

  const handlePageSizeChange = async (size) => {
    localStorage.setItem('page-size', size + '');
    setPageSize(size);
    setActivePage(1);
    loadLogs(activePage, size)
      .then()
      .catch((reason) => {
        showError(reason);
      });
  };

  // Refresh function
  const refresh = async () => {
    setActivePage(1);
    handleEyeClick();
    await loadLogs(1, pageSize);
  };

  // Copy text function
  const copyText = async (e, text) => {
    e.stopPropagation();
    if (await copy(text)) {
      showSuccess('已复制：' + text);
    } else {
      Modal.error({ title: t('无法复制到剪贴板，请手动复制'), content: text });
    }
  };

  // Initialize data
  useEffect(() => {
    const localPageSize =
      parseInt(localStorage.getItem('page-size')) || ITEMS_PER_PAGE;
    setPageSize(localPageSize);
    loadLogs(activePage, localPageSize)
      .then()
      .catch((reason) => {
        showError(reason);
      });
  }, []);

  // Initialize statistics when formApi is available
  useEffect(() => {
    if (formApi) {
      handleEyeClick();
    }
  }, [formApi]);

  // Check if any record has expandable content
  const hasExpandableRows = () => {
    return logs.some(
      (log) => expandData[log.key] && expandData[log.key].length > 0,
    );
  };

  return {
    // Basic state
    logs,
    expandData,
    showStat,
    loading,
    loadingStat,
    activePage,
    logCount,
    pageSize,
    logType,
    stat,
    isAdminUser,

    // Form state
    formApi,
    setFormApi,
    formInitValues,
    getFormValues,

    // Column visibility
    visibleColumns,
    showColumnSelector,
    setShowColumnSelector,
    handleColumnVisibilityChange,
    handleSelectAll,
    initDefaultColumns,
    COLUMN_KEYS,

    // Compact mode
    compactMode,
    setCompactMode,

    // User info modal
    showUserInfo,
    setShowUserInfoModal,
    userInfoData,
    showUserInfoFunc,

    // Functions
    loadLogs,
    handlePageChange,
    handlePageSizeChange,
    refresh,
    copyText,
    handleEyeClick,
    setLogsFormat,
    hasExpandableRows,
    setLogType,

    exportAllLogs,
    cancelExport,
    exportProgress,
    isExporting,
    // Translation
    t,
  };
};
