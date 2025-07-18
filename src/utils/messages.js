/**
 * messages.js
 * @description :: exports all response for APIS.
 */

const responseCode = require('./responseCode');

/**
 * @description: exports response format of all APIS
 * @param {obj | Array} data : object which will returned in response.
 * @param {obj} res : response from controller method.
 * @return {obj} : response for API {status, message, data}
 */
module.exports = {
  successResponse: (data, res) =>
    res.status(responseCode.success).json({
      status: 'SUCCESS',
      message: data.message || 'Your request is successfully executed',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  noContent: (data, res) =>
    res.status(responseCode.noContent).json({
      status: 'SUCCESS',
      message: data.message || 'No Content',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  failureResponse: (data, res) =>
    res.status(responseCode.internalServerError).json({
      status: 'FAILURE',
      message: data.message || 'Internal server error.',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  badRequest: (data, res) =>
    res.status(responseCode.badRequest).json({
      status: 'BAD_REQUEST',
      message:
        data.message || 'The request cannot be fulfilled due to bad syntax.',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  insufficientParameters: (data, res) =>
    res.status(responseCode.badRequest).json({
      status: 'BAD_REQUEST',
      message: data.message || 'Insufficient parameters.',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  unAuthorizedRequest: (data, res) =>
    res.status(responseCode.unAuthorizedRequest).json({
      status: 'UNAUTHORIZED',
      message: data.message || 'You are not authorized to access the request.',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  accessForbidden: (data, res) =>
    res.status(responseCode.accessForbidden).json({
      status: 'ACCESS_FORBIDEN',
      message: data.message || 'You are not authorized to access the request.',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  notFound: (data, res) =>
    res.status(responseCode.notFound).json({
      status: 'NOT FOUND',
      message: data.message || 'Data Not Found..',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  unProcessable_entity: (data, res) =>
    res.status(responseCode.unProcessable_entity).json({
      status: 'unProcessable_entity',
      message: data.message || 'unProcessable_entity..',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  badGateway: (data, res) =>
    res.status(responseCode.badGateway).json({
      status: 'badGateway',
      message: data.message || 'badGateway..',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),

  payment_required: (data, res) =>
    res.status(responseCode.payment_required).json({
      status: 'FAILURE',
      message: data.message || 'please provide payment..',
      data: data.data && Object.keys(data.data).length ? data.data : [],
    }),
};
