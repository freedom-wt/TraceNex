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

import React, { useEffect, useState, useRef, useMemo } from 'react';
import { Button, Col, Form, Row, Spin, Table, Input, Select, Space } from '@douyinfe/semi-ui';
import {
  compareObjects,
  API,
  showError,
  showSuccess,
  showWarning,
  verifyJSON,
} from '../../../helpers';
import { useTranslation } from 'react-i18next';

// 工具函数：GroupRatio JSON 转表格数据
const parseGroupRatioJSON = (jsonStr) => {
  if (!jsonStr) return [{ groupName: '', ratio: '' }];
  try {
    const obj = JSON.parse(jsonStr);
    return Object.entries(obj).map(([groupName, ratio]) => ({ groupName, ratio: String(ratio) }));
  } catch {
    return [{ groupName: '', ratio: '' }];
  }
};

// 工具函数：GroupRatio 表格数据转 JSON
const stringifyGroupRatioData = (data) => {
  const validData = data.filter(item => item.groupName.trim() && item.ratio.trim());
  const obj = validData.reduce((acc, item) => {
    acc[item.groupName.trim()] = Number(item.ratio);
    return acc;
  }, {});
  return JSON.stringify(obj);
};

// 工具函数：UserUsableGroups JSON 转表格数据
const parseUserUsableGroupsJSON = (jsonStr) => {
  if (!jsonStr) return [{ groupName: '', desc: '' }];
  try {
    const obj = JSON.parse(jsonStr);
    return Object.entries(obj).map(([groupName, desc]) => ({ groupName, desc: String(desc) }));
  } catch {
    return [{ groupName: '', desc: '' }];
  }
};

// 工具函数：UserUsableGroups 表格数据转 JSON
const stringifyUserUsableGroupsData = (data) => {
  const validData = data.filter(item => item.groupName.trim() && item.desc.trim());
  const obj = validData.reduce((acc, item) => {
    acc[item.groupName.trim()] = item.desc.trim();
    return acc;
  }, {});
  return JSON.stringify(obj);
};

// 工具函数：GroupGroupRatio JSON 转表格数据
const parseGroupGroupRatioJSON = (jsonStr) => {
  if (!jsonStr) return [{ outerGroup: '', innerGroup: '', ratio: '' }];
  try {
    const obj = JSON.parse(jsonStr);
    const rows = [];
    Object.entries(obj).forEach(([outerGroup, innerObj]) => {
      Object.entries(innerObj).forEach(([innerGroup, ratio]) => {
        rows.push({ outerGroup, innerGroup, ratio: String(ratio) });
      });
    });
    return rows.length ? rows : [{ outerGroup: '', innerGroup: '', ratio: '' }];
  } catch {
    return [{ outerGroup: '', innerGroup: '', ratio: '' }];
  }
};

// 工具函数：GroupGroupRatio 表格数据转 JSON
const stringifyGroupGroupRatioData = (data) => {
  const validData = data.filter(item =>
    item.outerGroup.trim() && item.innerGroup.trim() && item.ratio.trim()
  );
  const obj = validData.reduce((acc, item) => {
    const outer = item.outerGroup.trim();
    const inner = item.innerGroup.trim();
    const ratio = Number(item.ratio);
    if (!acc[outer]) acc[outer] = {};
    acc[outer][inner] = ratio;
    return acc;
  }, {});
  return JSON.stringify(obj);
};

// 工具函数：group_special_usable_group JSON 转表格数据
const parseSpecialUsableGroupJSON = (jsonStr) => {
  if (!jsonStr) return [{ userGroup: '', op: '+:', targetGroup: '', desc: '' }];
  try {
    const obj = JSON.parse(jsonStr);
    const rows = [];
    Object.entries(obj).forEach(([userGroup, opObj]) => {
      Object.entries(opObj).forEach(([key, value]) => {
        let op = '+:';
        let targetGroup = key;
        if (key.startsWith('+:')) {
          targetGroup = key.slice(2);
          op = '+:';
        } else if (key.startsWith('-:')) {
          targetGroup = key.slice(2);
          op = '-:';
        }
        rows.push({ userGroup, op, targetGroup, desc: String(value) });
      });
    });
    return rows.length ? rows : [{ userGroup: '', op: '+:', targetGroup: '', desc: '' }];
  } catch {
    return [{ userGroup: '', op: '+:', targetGroup: '', desc: '' }];
  }
};

const stringifySpecialUsableGroupData = (data) => {
  const validData = data.filter(item =>
    item.userGroup.trim() && item.op && item.targetGroup.trim() && item.desc.trim()
  );
  const obj = validData.reduce((acc, item) => {
    const userGroup = item.userGroup.trim();
    const op = item.op;
    const targetGroup = item.targetGroup.trim();
    const desc = item.desc.trim();
    // 修复：-: 操作的 key 拼接逻辑
    const key = op === '+:' ? `+:${targetGroup}` : `-:${targetGroup}`;

    if (!acc[userGroup]) acc[userGroup] = {};
    acc[userGroup][key] = desc;
    return acc;
  }, {});
  return JSON.stringify(obj);
};

// 工具函数：AutoGroups JSON 转表格数据
const parseAutoGroupsJSON = (jsonStr) => {
  if (!jsonStr) return [{ groupName: '' }];
  try {
    const arr = JSON.parse(jsonStr);
    return Array.isArray(arr) && arr.length
      ? arr.map(item => ({ groupName: String(item) }))
      : [{ groupName: '' }];
  } catch {
    return [{ groupName: '' }];
  }
};

// 工具函数：AutoGroups 表格数据转 JSON
const stringifyAutoGroupsData = (data) => {
  const validData = data.filter(item => item.groupName.trim()).map(item => item.groupName.trim());
  return JSON.stringify(validData);
};

// 工具函数：计算表格合并配置的工具函数（只要用户分组相同就合并，无视内容是否完整）
const getMergeSpanConfig = (data) => {
  if (!data || data.length === 0) return [];

  const spanConfig = [];
  const groupCountMap = {}; // 统计每个用户分组的总行数（不管是否填完）
  const groupUsedMap = {};  // 标记分组是否已处理

  // 步骤1：统计每个用户分组的总行数（所有行，不管内容是否完整）
  data.forEach(item => {
    const groupKey = item.userGroup.trim() || '__empty_group__'; // 空分组用固定标识
    groupCountMap[groupKey] = (groupCountMap[groupKey] || 0) + 1;
  });

  // 步骤2：为每个行生成合并配置
  data.forEach(item => {
    const groupKey = item.userGroup.trim() || '__empty_group__';
    if (!groupUsedMap[groupKey]) {
      // 该分组第一行：userGroup列合并，占满该分组的总行数
      spanConfig.push({ rowSpan: groupCountMap[groupKey], hidden: false });
      groupUsedMap[groupKey] = true;
    } else {
      // 该分组其他行：userGroup列隐藏（合并到第一行）
      spanConfig.push({ rowSpan: 0, hidden: true });
    }
  });
  return spanConfig;
};

export default function GroupRatioSettings(props) {
  const { t } = useTranslation();
  const [loading, setLoading] = useState(false);
  const refForm = useRef();

  // 表格数据状态
  const [groupRatioData, setGroupRatioData] = useState([{ groupName: '', ratio: '' }]);
  const [userUsableGroupsData, setUserUsableGroupsData] = useState([{ groupName: '', desc: '' }]);
  const [groupGroupRatioData, setGroupGroupRatioData] = useState([{ outerGroup: '', innerGroup: '', ratio: '' }]);
  const [specialUsableGroupData, setSpecialUsableGroupData] = useState([{ userGroup: '', op: '+:', targetGroup: '', desc: '' }]);
  const [autoGroupsData, setAutoGroupsData] = useState([{ groupName: '' }]);
  const [defaultUseAutoGroup, setDefaultUseAutoGroup] = useState(false);
  // 新增：合并配置的state（依赖specialUsableGroupData变化）
  const [mergeSpanConfig, setMergeSpanConfig] = useState([]);
  // 新增：记录选中行索引（用于新增行继承分组）
  const [selectedRowIndex, setSelectedRowIndex] = useState(-1); // 新增这行

  // 新增：监听表格数据变化，更新合并配置
  useEffect(() => {
    const config = getMergeSpanConfig(specialUsableGroupData);
    setMergeSpanConfig(config);
  }, [specialUsableGroupData]);

  // 提交处理
  async function onSubmit() {
    try {
      // 表格数据转 JSON 字符串
      const formData = {
        GroupRatio: stringifyGroupRatioData(groupRatioData),
        UserUsableGroups: stringifyUserUsableGroupsData(userUsableGroupsData),
        GroupGroupRatio: stringifyGroupGroupRatioData(groupGroupRatioData),
        'group_ratio_setting.group_special_usable_group': stringifySpecialUsableGroupData(specialUsableGroupData),
        AutoGroups: stringifyAutoGroupsData(autoGroupsData),
        DefaultUseAutoGroup: defaultUseAutoGroup,
      };

      // JSON 验证
      const validateResult = Object.entries(formData).every(([key, value]) => {
        if (key === 'DefaultUseAutoGroup') return true;
        return verifyJSON(value);
      });

      if (!validateResult) {
        showError(t('表格数据转换后不是合法的 JSON 字符串'));
        return;
      }

      // 对比原始数据
      const originalData = {
        GroupRatio: props.options.GroupRatio || '',
        UserUsableGroups: props.options.UserUsableGroups || '',
        GroupGroupRatio: props.options.GroupGroupRatio || '',
        'group_ratio_setting.group_special_usable_group': props.options['group_ratio_setting.group_special_usable_group'] || '',
        AutoGroups: props.options.AutoGroups || '',
        DefaultUseAutoGroup: props.options.DefaultUseAutoGroup || false,
      };

      const updateArray = compareObjects(formData, originalData);
      if (!updateArray.length) {
        return showWarning(t('你似乎并没有修改什么'));
      }

      // 构建请求队列
      const requestQueue = updateArray.map((item) => {
        const value = typeof formData[item.key] === 'boolean'
          ? String(formData[item.key])
          : formData[item.key];
        return API.put('/api/option/', { key: item.key, value });
      });

      setLoading(true);
      Promise.all(requestQueue)
        .then((res) => {
          if (res.includes(undefined)) {
            return showError(
              requestQueue.length > 1
                ? t('部分保存失败，请重试')
                : t('保存失败'),
            );
          }

          for (let i = 0; i < res.length; i++) {
            if (!res[i].data.success) {
              return showError(res[i].data.message);
            }
          }

          showSuccess(t('保存成功'));
          props.refresh();
        })
        .catch((error) => {
          console.error('Unexpected error:', error);
          showError(t('保存失败，请重试'));
        })
        .finally(() => {
          setLoading(false);
        });
    } catch (error) {
      showError(t('请检查输入'));
      console.error(error);
      setLoading(false);
    }
  }

  // 初始化数据
  useEffect(() => {
    if (!props.options) return;
    // 解析各字段为表格数据
    setGroupRatioData(parseGroupRatioJSON(props.options.GroupRatio || ''));
    setUserUsableGroupsData(parseUserUsableGroupsJSON(props.options.UserUsableGroups || ''));
    setGroupGroupRatioData(parseGroupGroupRatioJSON(props.options.GroupGroupRatio || ''));
    setSpecialUsableGroupData(parseSpecialUsableGroupJSON(props.options['group_ratio_setting.group_special_usable_group'] || ''));
    setAutoGroupsData(parseAutoGroupsJSON(props.options.AutoGroups || ''));
    setDefaultUseAutoGroup(props.options.DefaultUseAutoGroup || false);
  }, [props.options]);

  // 表格行操作：新增行（优化：支持继承选中行的分组）
  const addRow = (setter, defaultRow, inheritGroup = '') => { // 新增inheritGroup参数
    setter(prev => {
      // 继承选中行的分组（如果有）
      const newRow = inheritGroup ? { ...defaultRow, userGroup: inheritGroup } : defaultRow;
      return [...prev, newRow];
    });
  };

  // 表格行操作：删除行
  const deleteRow = (setter, index) => {
    setter(prev => prev.filter((_, i) => i !== index));
  };

  // 分组倍率表格列配置
  const groupRatioColumns = [
    {
      title: t('分组名称'),
      dataIndex: 'groupName',
      render: (text, record, index) => (
        <Input
          value={text}
          onChange={(value) => {
            const newData = [...groupRatioData];
            newData[index].groupName = value;
            setGroupRatioData(newData);
          }}
          placeholder={t('请输入分组名称')}
        />
      ),
    },
    {
      title: t('倍率'),
      dataIndex: 'ratio',
      render: (text, record, index) => (
        <Input
          value={text}
          onChange={(value) => {
            const newData = [...groupRatioData];
            newData[index].ratio = value;
            setGroupRatioData(newData);
          }}
          placeholder={t('请输入数字倍率')}
          type="number"
          step="0.1"
        />
      ),
    },
    {
      title: t('操作'),
      render: (_, record, index) => (
        <Space>
          <Button
            theme="danger"
            size="small"
            onClick={() => deleteRow(setGroupRatioData, index)}
            disabled={groupRatioData.length === 1}
          >
            {t('删除')}
          </Button>
        </Space>
      ),
    },
  ];

  // 用户可选分组表格列配置
  const userUsableGroupsColumns = [
    {
      title: t('分组名称'),
      dataIndex: 'groupName',
      render: (text, record, index) => (
        <Input
          value={text}
          onChange={(value) => {
            const newData = [...userUsableGroupsData];
            newData[index].groupName = value;
            setUserUsableGroupsData(newData);
          }}
          placeholder={t('请输入分组名称')}
        />
      ),
    },
    {
      title: t('分组描述'),
      dataIndex: 'desc',
      render: (text, record, index) => (
        <Input
          value={text}
          onChange={(value) => {
            const newData = [...userUsableGroupsData];
            newData[index].desc = value;
            setUserUsableGroupsData(newData);
          }}
          placeholder={t('请输入分组描述')}
        />
      ),
    },
    {
      title: t('操作'),
      render: (_, record, index) => (
        <Space>
          <Button
            theme="danger"
            size="small"
            onClick={() => deleteRow(setUserUsableGroupsData, index)}
            disabled={userUsableGroupsData.length === 1}
          >
            {t('删除')}
          </Button>
        </Space>
      ),
    },
  ];

  // 分组特殊倍率表格列配置
  const groupGroupRatioColumns = [
    {
      title: t('外层分组'),
      dataIndex: 'outerGroup',
      render: (text, record, index) => (
        <Input
          value={text}
          onChange={(value) => {
            const newData = [...groupGroupRatioData];
            newData[index].outerGroup = value;
            setGroupGroupRatioData(newData);
          }}
          placeholder={t('请输入外层分组名称')}
        />
      ),
    },
    {
      title: t('内层分组'),
      dataIndex: 'innerGroup',
      render: (text, record, index) => (
        <Input
          value={text}
          onChange={(value) => {
            const newData = [...groupGroupRatioData];
            newData[index].innerGroup = value;
            setGroupGroupRatioData(newData);
          }}
          placeholder={t('请输入内层分组名称')}
        />
      ),
    },
    {
      title: t('倍率'),
      dataIndex: 'ratio',
      render: (text, record, index) => (
        <Input
          value={text}
          onChange={(value) => {
            const newData = [...groupGroupRatioData];
            newData[index].ratio = value;
            setGroupGroupRatioData(newData);
          }}
          placeholder={t('请输入数字倍率')}
          type="number"
          step="0.1"
        />
      ),
    },
    {
      title: t('操作'),
      render: (_, record, index) => (
        <Space>
          <Button
            theme="danger"
            size="small"
            onClick={() => deleteRow(setGroupGroupRatioData, index)}
            disabled={groupGroupRatioData.length === 1}
          >
            {t('删除')}
          </Button>
        </Space>
      ),
    },
  ];
  const addItemToSameUserGroupEnd = (originalArray, newItem) => {
    const targetUserGroup = newItem.userGroup?.trim() || '';
    const newArray = [...originalArray];

    // 找到最后一个同userGroup的项的索引
    let lastSameGroupIndex = -1;
    newArray.forEach((item, index) => {
      if (item.userGroup?.trim() === targetUserGroup) {
        lastSameGroupIndex = index;
      }
    });

    // 插入逻辑：存在同分组项则插入到其后方，否则追加到末尾
    if (lastSameGroupIndex !== -1) {
      newArray.splice(lastSameGroupIndex + 1, 0, newItem);
    } else {
      newArray.push(newItem);
    }
    return newArray;
  };

  const onChangeValue = (newValue, record, type, index) => {
    const trimmedNewValue = newValue.trim();
    const oldValue = record[type].trim();
    // 函数式更新，确保拿到最新的状态
    if (type === 'useGroup') {
      setSpecialUsableGroupData(prevData => {
        const list = prevData.map(item => {
          if (item[type].trim() === oldValue) {
            return { ...item, userGroup: trimmedNewValue };
          }
          return item;
        });
        const newList = list.filter((item, ind) => ind !== index);
        const newObj = { ...record, userGroup: trimmedNewValue };
        const newDta = addItemToSameUserGroupEnd(newList, newObj);
        return newDta;
      });
    } else {
      setSpecialUsableGroupData(prevData => {
        return prevData.map(item => {
          if (item[type].trim() === oldValue) {
            return { ...item, [type]: trimmedNewValue }
          }
          return item
        })
      })
    }
  };
  // 分组特殊可用分组表格列配置（添加行列合并逻辑）
  const specialUsableGroupColumns = useMemo(() => {
    return [
      {
        title: t('用户分组'),
        dataIndex: 'userGroup',
        width: 120,
        // 关键：自定义单元格渲染，实现行合并
        render: (value, record, index) => {
          const { rowSpan, hidden } = mergeSpanConfig[index] || { rowSpan: 1, hidden: false };
          // 修复：userGroup 同步逻辑（先取新值，再匹配旧分组）
          return {
            children: hidden ? null : ( // 隐藏的单元格不渲染内容
              <Input
                value={value}
                onChange={(newValue) => {
                  onChangeValue(newValue, record, 'userGroup', index);
                }}
                placeholder={t('请输入用户分组名称（如default/黄金）')}
                style={{ width: '100%' }}
              />
            ),
            props: {
              rowSpan, // 设置行合并数
              style: hidden ? { display: 'none' } : {}, // 隐藏多余单元格
            },
          };
        },
      },
      {
        title: t('操作类型'),
        dataIndex: 'op',
        width: 120,
        render: (text, record, index) => (
          <Select
            value={text}
            onChange={(value) => {
              const newData = [...specialUsableGroupData];
              newData[index].op = value;
              setSpecialUsableGroupData(newData);
            }}
            optionList={[
              { label: t('添加分组'), value: '+:' },
              { label: t('移除分组'), value: '-:' },
            ]}
            style={{ width: 160 }}
            placeholder={t('选择操作类型')}
          />
        ),
      },
      {
        title: t('目标分组'),
        dataIndex: 'targetGroup',
        width: 150,
        render: (text, record, index) => (
          <Input
            value={text}
            onChange={(newValue) => {
              onChangeValue(newValue, record, 'targetGroup', index);
            }}
            placeholder={t('请输入目标分组名（如黄金/append_1）')}
          />
        ),
      },
      {
        title: t('描述/值'),
        dataIndex: 'desc',
        width: 180,
        render: (text, record, index) => (
          <Input
            value={text}
            onChange={(newValue) => {
              onChangeValue(newValue, record, 'desc', index);
            }}
            placeholder={t('请输入描述信息（如黄金分组）')}
          />
        ),
      },
      {
        title: t('操作'),
        width: 80,
        render: (_, record, index) => (
          <Space>
            <Button
              theme="danger"
              size="small"
              onClick={() => {
                // 删除行后，合并配置会自动重新计算
                deleteRow(setSpecialUsableGroupData, index);
              }}
              disabled={specialUsableGroupData.length === 1}
            >
              {t('删除')}
            </Button>
          </Space>
        ),
      },
    ];
  }, [mergeSpanConfig, t])
  // 自动分组表格列配置
  const autoGroupsColumns = [
    {
      title: t('分组名称'),
      dataIndex: 'groupName',
      render: (text, record, index) => (
        <Input
          value={text}
          onChange={(value) => {
            const newData = [...autoGroupsData];
            newData[index].groupName = value;
            setAutoGroupsData(newData);
          }}
          placeholder={t('请输入自动分组名称')}
        />
      ),
    },
    {
      title: t('操作'),
      render: (_, record, index) => (
        <Space>
          <Button
            theme="danger"
            size="small"
            onClick={() => deleteRow(setAutoGroupsData, index)}
            disabled={autoGroupsData.length === 1}
          >
            {t('删除')}
          </Button>
        </Space>
      ),
    },
  ];

  return (
    <Spin spinning={loading}>
      <Form
        getFormApi={(formAPI) => (refForm.current = formAPI)}
        style={{ marginBottom: 15 }}
      >
        {/* 分组倍率 */}
        <Row gutter={16} style={{ marginBottom: 20 }}>
          <Col xs={24}>
            <h4>{t('分组倍率')}</h4>
            <p style={{ color: '#666', marginBottom: 10 }}>
              {t('分组倍率设置，可以在此处新增分组或修改现有分组的倍率')}
            </p>
            <Table
              columns={groupRatioColumns}
              dataSource={groupRatioData}
              pagination={false}
              bordered
            />
            <div className='flex justify-between'>
              <Button
                theme="primary"
                size="small"
                style={{ marginTop: 10 }}
                onClick={() => addRow(setGroupRatioData, { groupName: '', ratio: '' })}
              >
                {t('新增行')}
              </Button>
              <p style={{ color: '#666' }} className='mt-2'>
                {t('总计')}
                <span style={{ color: 'rgb(0, 100, 250)', fontWeight: 600 }} className='mr-2 ml-2'>
                  {groupRatioData ? groupRatioData.length : 0}
                </span>
                {t('条')}
              </p></div>
          </Col>
        </Row>

        {/* 用户可选分组 */}
        <Row gutter={16} style={{ marginBottom: 20 }}>
          <Col xs={24}>
            <h4>{t('用户可选分组')}</h4>
            <p style={{ color: '#666', marginBottom: 10 }}>
              {t('用户新建令牌时可选的分组')}
            </p>
            <Table
              columns={userUsableGroupsColumns}
              dataSource={userUsableGroupsData}
              pagination={false}
              bordered
            />
            <div className='flex justify-between'>
              <Button
                theme="primary"
                size="small"
                style={{ marginTop: 10 }}
                onClick={() => addRow(setUserUsableGroupsData, { groupName: '', desc: '' })}
              >
                {t('新增行')}
              </Button>
              <p style={{ color: '#666' }} className='mt-2'>
                {t('总计')}
                <span style={{ color: 'rgb(0, 100, 250)', fontWeight: 600 }} className='mr-2 ml-2'>
                  {userUsableGroupsData ? userUsableGroupsData.length : 0}
                </span>
                {t('条')}
              </p></div>
          </Col>
        </Row>

        {/* 分组特殊倍率 */}
        <Row gutter={16} style={{ marginBottom: 20 }}>
          <Col xs={24}>
            <h4>{t('分组特殊倍率')}</h4>
            <p style={{ color: '#666', marginBottom: 10 }}>
              {t('键为分组名称，值为另一个 JSON 对象，键为分组名称，值为该分组的用户的特殊分组倍率')}
            </p>
            <Table
              columns={groupGroupRatioColumns}
              dataSource={groupGroupRatioData}
              pagination={false}
              bordered
            />
            <div className='flex justify-between'>
              <Button
                theme="primary"
                size="small"
                style={{ marginTop: 10 }}
                onClick={() => addRow(setGroupGroupRatioData, { outerGroup: '', innerGroup: '', ratio: '' })}
              >
                {t('新增行')}
              </Button>
              <p style={{ color: '#666' }} className='mt-2'>
                {t('总计')}
                <span style={{ color: 'rgb(0, 100, 250)', fontWeight: 600 }} className='mr-2 ml-2'>
                  {groupGroupRatioData ? groupGroupRatioData.length : 0}
                </span>
                {t('条')}
              </p></div>
          </Col>
        </Row>

        {/* 分组特殊可用分组 */}
        <Row gutter={16} style={{ marginBottom: 20 }}>
          <Col xs={24}>
            <h4>{t('分组特殊可用分组')}</h4>
            <p style={{ color: '#666', marginBottom: 10 }}>
              {t('键为用户分组名称，值为操作映射对象。内层键以"+:"开头表示添加指定分组，以"-:"开头表示移除指定分组')}
            </p>
            <Table
              columns={specialUsableGroupColumns}
              dataSource={specialUsableGroupData}
              pagination={false}
              bordered
              scroll={{ x: '100%', y: '500px' }}
            />
            <div className='flex justify-between'>
              <Button
                theme="primary"
                size="small"
                style={{ marginTop: 10 }}
                onClick={() => {
                  const inheritGroup = selectedRowIndex >= 0
                    ? specialUsableGroupData[selectedRowIndex].userGroup
                    : '';
                  addRow(
                    setSpecialUsableGroupData,
                    { userGroup: '', op: '+:', targetGroup: '', desc: '' },
                    inheritGroup
                  );
                }}
              >
                {t('新增行')}
              </Button>
              <p style={{ color: '#666' }} className='mt-2'>
                {t('总计')}
                <span style={{ color: 'rgb(0, 100, 250)', fontWeight: 600 }} className='mr-2 ml-2'>
                  {specialUsableGroupData ? specialUsableGroupData.length : 0}
                </span>
                {t('条')}
              </p>
            </div>

          </Col>
        </Row>

        {/* 自动分组 */}
        <Row gutter={16} style={{ marginBottom: 20 }}>
          <Col xs={24}>
            <h4>{t('自动分组auto，从第一个开始选择')}</h4>
            <p style={{ color: '#666', marginBottom: 10 }}>
              {t('自动分组列表，按顺序匹配')}
            </p>
            <Table
              columns={autoGroupsColumns}
              dataSource={autoGroupsData}
              pagination={false}
              bordered
            />
            <Button
              theme="primary"
              size="small"
              style={{ marginTop: 10 }}
              onClick={() => addRow(setAutoGroupsData, { groupName: '' })}
            >
              {t('新增行')}
            </Button>
          </Col>
        </Row>

        {/* 默认选择auto分组开关 */}
        <Row gutter={16} style={{ marginBottom: 20 }}>
          <Col xs={24} sm={16}>
            <Form.Switch
              label={t(
                '创建令牌默认选择auto分组，初始令牌也将设为auto（否则留空，为用户默认分组）',
              )}
              value={defaultUseAutoGroup}
              onChange={(value) => setDefaultUseAutoGroup(value)}
            />
          </Col>
        </Row>
      </Form>

      <Button onClick={onSubmit} type="primary">
        {t('保存分组倍率设置')}
      </Button>
    </Spin>
  );
}