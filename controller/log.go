package controller

import (
	"encoding/csv"
	"net/http"
	"strconv"
	"time"

	"github.com/QuantumNous/new-api/common"
	"github.com/QuantumNous/new-api/model"

	"github.com/gin-gonic/gin"
)

func GetAllLogs(c *gin.Context) {
	pageInfo := common.GetPageQuery(c)
	logType, _ := strconv.Atoi(c.Query("type"))
	startTimestamp, _ := strconv.ParseInt(c.Query("start_timestamp"), 10, 64)
	endTimestamp, _ := strconv.ParseInt(c.Query("end_timestamp"), 10, 64)
	username := c.Query("username")
	tokenName := c.Query("token_name")
	modelName := c.Query("model_name")
	channel, _ := strconv.Atoi(c.Query("channel"))
	group := c.Query("group")
	logs, total, err := model.GetAllLogs(logType, startTimestamp, endTimestamp, modelName, username, tokenName, pageInfo.GetStartIdx(), pageInfo.GetPageSize(), channel, group)
	if err != nil {
		common.ApiError(c, err)
		return
	}
	pageInfo.SetTotal(int(total))
	pageInfo.SetItems(logs)
	common.ApiSuccess(c, pageInfo)
	return
}

func GetUserLogs(c *gin.Context) {
	pageInfo := common.GetPageQuery(c)
	userId := c.GetInt("id")
	logType, _ := strconv.Atoi(c.Query("type"))
	startTimestamp, _ := strconv.ParseInt(c.Query("start_timestamp"), 10, 64)
	endTimestamp, _ := strconv.ParseInt(c.Query("end_timestamp"), 10, 64)
	tokenName := c.Query("token_name")
	modelName := c.Query("model_name")
	group := c.Query("group")
	logs, total, err := model.GetUserLogs(userId, logType, startTimestamp, endTimestamp, modelName, tokenName, pageInfo.GetStartIdx(), pageInfo.GetPageSize(), group)
	if err != nil {
		common.ApiError(c, err)
		return
	}
	pageInfo.SetTotal(int(total))
	pageInfo.SetItems(logs)
	common.ApiSuccess(c, pageInfo)
	return
}

// writeLogsCSV 将日志列表直接流式写入 HTTP 响应为 CSV 文件
func writeLogsCSV(c *gin.Context, logs []*model.Log) {
	// 设置下载头，先发送给客户端，方便尽早开始下载
	c.Header("Content-Disposition", "attachment; filename=logs.csv")
	c.Header("Content-Type", "text/csv; charset=utf-8")

	// 写入 UTF-8 BOM，便于 Excel 正确识别编码
	_, _ = c.Writer.Write([]byte("\xEF\xBB\xBF"))

	w := csv.NewWriter(c.Writer)
	// 表头与前端展示表格含义对应（顺序：时间、渠道、用户、令牌、分组、类型、模型、用时/首字、输入、输出、花费、IP、重试、详情）
	header := []string{"时间", "渠道", "用户", "令牌", "分组", "类型", "模型", "用时/首字", "输入", "输出", "花费", "IP", "重试", "详情"}
	if err := w.Write(header); err != nil {
		// 这里已经开始写出响应，只能尽量结束 writer
		w.Flush()
		return
	}
	for _, log := range logs {
		// 时间
		timeStr := ""
		if log.CreatedAt != 0 {
			timeStr = time.Unix(log.CreatedAt, 0).Format("2006-01-02 15:04:05")
		}
		// 渠道（与前端“渠道”列含义一致：ID + 名称）
		channelStr := ""
		if log.ChannelId != 0 || log.ChannelName != "" {
			channelStr = strconv.Itoa(log.ChannelId)
			if log.ChannelName != "" {
				channelStr = channelStr + " - " + log.ChannelName
			}
		}
		// 用时/首字：这里导出总用时（秒），首字耗时前端是从 other 里计算的，CSV 里先只保留总用时
		useTimeStr := ""
		if log.UseTime > 0 {
			useTimeStr = strconv.Itoa(log.UseTime)
		}
		// 输入 / 输出 / 花费
		promptStr := ""
		if log.PromptTokens > 0 {
			promptStr = strconv.Itoa(log.PromptTokens)
		}
		completionStr := ""
		if log.CompletionTokens > 0 {
			completionStr = strconv.Itoa(log.CompletionTokens)
		}
		quotaStr := ""
		if log.Quota != 0 {
			quotaStr = strconv.Itoa(log.Quota)
		}

		row := []string{
			timeStr,          // 时间
			channelStr,       // 渠道
			log.Username,     // 用户
			log.TokenName,    // 令牌
			log.Group,        // 分组
			strconv.Itoa(log.Type), // 类型
			log.ModelName,    // 模型
			useTimeStr,       // 用时/首字（总用时秒）
			promptStr,        // 输入
			completionStr,    // 输出
			quotaStr,         // 花费
			log.Ip,           // IP
			"",               // 重试（前端从 other 计算，这里暂留空）
			log.Content,      // 详情（使用 content 字段）
		}
		_ = w.Write(row)
	}
	w.Flush()
}

// ExportAllLogs 管理员导出全部日志（当前筛选条件下，最多 MaxLogExportItems 条），返回 CSV
func ExportAllLogs(c *gin.Context) {
	logType, _ := strconv.Atoi(c.Query("type"))
	startTimestamp, _ := strconv.ParseInt(c.Query("start_timestamp"), 10, 64)
	endTimestamp, _ := strconv.ParseInt(c.Query("end_timestamp"), 10, 64)
	username := c.Query("username")
	tokenName := c.Query("token_name")
	modelName := c.Query("model_name")
	channel, _ := strconv.Atoi(c.Query("channel"))
	group := c.Query("group")
	logs, err := model.GetAllLogsForExport(logType, startTimestamp, endTimestamp, modelName, username, tokenName, channel, group)
	if err != nil {
		common.ApiError(c, err)
		return
	}
	writeLogsCSV(c, logs)
}

// ExportUserLogs 用户导出自己的日志（当前筛选条件下，最多 MaxLogExportItems 条），返回 CSV
func ExportUserLogs(c *gin.Context) {
	userId := c.GetInt("id")
	logType, _ := strconv.Atoi(c.Query("type"))
	startTimestamp, _ := strconv.ParseInt(c.Query("start_timestamp"), 10, 64)
	endTimestamp, _ := strconv.ParseInt(c.Query("end_timestamp"), 10, 64)
	tokenName := c.Query("token_name")
	modelName := c.Query("model_name")
	group := c.Query("group")
	logs, err := model.GetUserLogsForExport(userId, logType, startTimestamp, endTimestamp, modelName, tokenName, group)
	if err != nil {
		common.ApiError(c, err)
		return
	}
	writeLogsCSV(c, logs)
}

func SearchAllLogs(c *gin.Context) {
	keyword := c.Query("keyword")
	logs, err := model.SearchAllLogs(keyword)
	if err != nil {
		common.ApiError(c, err)
		return
	}
	c.JSON(http.StatusOK, gin.H{
		"success": true,
		"message": "",
		"data":    logs,
	})
	return
}

func SearchUserLogs(c *gin.Context) {
	keyword := c.Query("keyword")
	userId := c.GetInt("id")
	logs, err := model.SearchUserLogs(userId, keyword)
	if err != nil {
		common.ApiError(c, err)
		return
	}
	c.JSON(http.StatusOK, gin.H{
		"success": true,
		"message": "",
		"data":    logs,
	})
	return
}

func GetLogByKey(c *gin.Context) {
	key := c.Query("key")
	logs, err := model.GetLogByKey(key)
	if err != nil {
		c.JSON(200, gin.H{
			"success": false,
			"message": err.Error(),
		})
		return
	}
	c.JSON(200, gin.H{
		"success": true,
		"message": "",
		"data":    logs,
	})
}

func GetLogsStat(c *gin.Context) {
	logType, _ := strconv.Atoi(c.Query("type"))
	startTimestamp, _ := strconv.ParseInt(c.Query("start_timestamp"), 10, 64)
	endTimestamp, _ := strconv.ParseInt(c.Query("end_timestamp"), 10, 64)
	tokenName := c.Query("token_name")
	username := c.Query("username")
	modelName := c.Query("model_name")
	channel, _ := strconv.Atoi(c.Query("channel"))
	group := c.Query("group")
	stat := model.SumUsedQuota(logType, startTimestamp, endTimestamp, modelName, username, tokenName, channel, group)
	//tokenNum := model.SumUsedToken(logType, startTimestamp, endTimestamp, modelName, username, "")
	c.JSON(http.StatusOK, gin.H{
		"success": true,
		"message": "",
		"data": gin.H{
			"quota": stat.Quota,
			"rpm":   stat.Rpm,
			"tpm":   stat.Tpm,
		},
	})
	return
}

func GetLogsSelfStat(c *gin.Context) {
	username := c.GetString("username")
	logType, _ := strconv.Atoi(c.Query("type"))
	startTimestamp, _ := strconv.ParseInt(c.Query("start_timestamp"), 10, 64)
	endTimestamp, _ := strconv.ParseInt(c.Query("end_timestamp"), 10, 64)
	tokenName := c.Query("token_name")
	modelName := c.Query("model_name")
	channel, _ := strconv.Atoi(c.Query("channel"))
	group := c.Query("group")
	quotaNum := model.SumUsedQuota(logType, startTimestamp, endTimestamp, modelName, username, tokenName, channel, group)
	//tokenNum := model.SumUsedToken(logType, startTimestamp, endTimestamp, modelName, username, tokenName)
	c.JSON(200, gin.H{
		"success": true,
		"message": "",
		"data": gin.H{
			"quota": quotaNum.Quota,
			"rpm":   quotaNum.Rpm,
			"tpm":   quotaNum.Tpm,
			//"token": tokenNum,
		},
	})
	return
}

func DeleteHistoryLogs(c *gin.Context) {
	targetTimestamp, _ := strconv.ParseInt(c.Query("target_timestamp"), 10, 64)
	if targetTimestamp == 0 {
		c.JSON(http.StatusOK, gin.H{
			"success": false,
			"message": "target timestamp is required",
		})
		return
	}
	count, err := model.DeleteOldLog(c.Request.Context(), targetTimestamp, 100)
	if err != nil {
		common.ApiError(c, err)
		return
	}
	c.JSON(http.StatusOK, gin.H{
		"success": true,
		"message": "",
		"data":    count,
	})
	return
}
