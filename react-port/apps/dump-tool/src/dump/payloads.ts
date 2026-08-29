import type { GraphQLRequest } from '@mygate/shared'
import { formatFilterDate, getMidnightEpoch } from '@mygate/shared'

export const EXPORT_REQUEST_QUERY = `query getAdminSrList($requestData: DataListInput) {
  getAdminSrList(requestData: $requestData) {
    dataResponse {
      data
      filterType
      totalCount
      __typename
    }
    message
    statusCode
    success
    __typename
  }
}
`

export const POLL_STATUS_QUERY = `query getDownloadReportList($requestData: DataListInput) {
  getDownloadReportList(requestData: $requestData) {
    dataResponse {
      data
      filterType
      totalCount
      __typename
    }
    message
    statusCode
    success
    __typename
  }
}
`

export interface ExportChunkInput {
  fromDate: string
  toDate: string
}

export interface DownloadReport {
  report_name?: string
  status?: string
  report_link?: string
  download_filters?: Record<string, string> | null
}

export function buildExportRequestPayload(chunk: ExportChunkInput): GraphQLRequest {
  const filterFrom = formatFilterDate(chunk.fromDate)
  const filterTo = formatFilterDate(chunk.toDate)

  return {
    operationName: 'getAdminSrList',
    variables: {
      requestData: {
        requiredFields: [
          'id',
          'number',
          'subject',
          'category',
          'sub_category',
          'house',
          'assignee',
          'mygate_status',
          'escalated_group',
          'defaulter',
          'updated_date',
          'urgent',
          'highlight_ticket',
          'ageing',
        ],
        pagination: { count: 25, page: 1 },
        sorting: [],
        conditions: [
          { name: 'date_filter', operation: 'equal', values: ['created_date'] },
          {
            name: 'mygate_status',
            values: ['open', 're_opened', 'in_progress', 'job_done', 'hold'],
            operation: 'in',
          },
          { name: 'from_date', values: [getMidnightEpoch(chunk.fromDate)], operation: 'gte' },
          { name: 'to_date', values: [getMidnightEpoch(chunk.toDate) + 86399], operation: 'lte' },
        ],
        isDownload: true,
        downloadFilters: {
          Status: ['New ', 'Reopened ', 'In Progress ', 'Job Done ', 'On Hold '],
          From: filterFrom,
          To: filterTo,
          'Date To': filterTo,
          'Date From': filterFrom,
        },
      },
    },
    query: EXPORT_REQUEST_QUERY,
  }
}

export function buildPollStatusPayload(): GraphQLRequest {
  return {
    operationName: 'getDownloadReportList',
    variables: {
      requestData: {
        requiredFields: [
          'report_link',
          'report_name',
          'status',
          'download_request_time',
          'download_filters',
        ],
        pagination: { count: 20, page: 1 },
        sorting: [],
        conditions: [],
      },
    },
    query: POLL_STATUS_QUERY,
  }
}
