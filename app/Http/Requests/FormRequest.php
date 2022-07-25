<?php
/**
 * Request : FormRequest.
 *
 * This file used for FormRequest to handle api request validation with json error
 *
 * @author JQ Gan <gan.jinqu@i-serve.com.my>
 */

namespace App\Http\Requests;

use Illuminate\Contracts\Validation\Validator;
use Illuminate\Http\Exceptions\HttpResponseException;
use Illuminate\Foundation\Http\FormRequest as CustomFormRequest;
use App\Traits\CipherTrait;
use App\Traits\ResponseTrait;
use Carbon\Carbon;

abstract class FormRequest extends CustomFormRequest
{
    use ResponseTrait;

    protected $responseErrors;
    protected $exceptionError;

    /**
     * Get the validation rules that apply to the request.
     *
     * @return array
     */
    abstract public function rules();

    /**
     * Determine if the user is authorized to make this request.
     *
     * @return bool
     */
    abstract public function authorize();

    /**
     * Handle a failed validation attempt.
     *
     * @param \Illuminate\Contracts\Validation\Validator $validator
     *
     * @throws \Illuminate\Validation\ValidationException
     */
    protected function failedValidation(Validator $validator)
    {
        $result = $this->requestErrors();

        $errors = [];
        $responses = null;
        $source = $this->header('source') ?? null;
        $validators = $validator->errors()->toArray();

        $failed = $validator->failed();

        foreach ($validators as $key => $value) {
            if ($failed[$key]) {
                $failedRule = array_key_first($failed[$key]);

                if (!empty($this->responseErrors[$failedRule])) {
                    $responses = $this->responseErrors[$failedRule];
                }
            }

            $errors[$key] = $value[0];
        }

        $result['message'] = !empty($errors) ? reset($errors) : NULL;
        $result['error'] = $errors;

        unset($errors);
        unset($this->responseErrors);

        if ($responses) {
            $result['data']['attributes'] = $responses;
        }

        throw new HttpResponseException(response()->json($result, $this->unprocessableStatus));
    }

    /**
     * Handle after validation rules success
     */
    public function withValidator($validator)
    {
        if (method_exists($this, 'withValidated')) {
            $this->withValidated($validator);
        }

        $sources = config('constant.source_authentication');

        if (!empty($sources[HEADER_SOURCE])) {
            $this->request->add(['source' => $sources[HEADER_SOURCE]]);
        }

        $this->replace($this->request->all());
    }

    /**
     * Prepare the data for validation.
     *
     * @return void
     */
    protected function prepareForValidation()
    {
        // $encryptedRequest = $this->cipherEncrypt(env('ENCRYPTION_KEY'), json_encode($this->all()));

        if (!empty($this->get('secret_encrypt'))) {
            $decryptedRequest = $this->cipherDecrypt(env('ENCRYPTION_KEY'), $this->get('secret_encrypt'));
            $requestData = json_decode($decryptedRequest, true);

            $this->request->remove('secret_encrypt');
            $this->request->add($requestData);
            $this->replace($this->request->all());
        }

        $cores = ['pagination'];

        if (method_exists($this, 'fillable')) {
            $this->replace($this->only(array_merge($this->fillable(), $cores)));
        } else {
            if (!empty($this->rules())) {
                $this->replace($this->only(array_merge(array_keys($this->rules()), $cores)));
            }
        }

        if (method_exists($this, 'validationFactory')) {
            $this->validationFactory();
        }
    }

    /**
     * Handle a failed authorization attempt.
     *
     * @return void
     *
     * @throws \Illuminate\Auth\Access\AuthorizationException
     */
    protected function failedAuthorization()
    {
        $result = $this->requestErrors();
        $result['error_code'] = 'FR001';
        $result['message'] = trans('general.unauthorized');
        $result['exception'] = $this->exceptionError;

        throw new HttpResponseException(response()->json($result, $this->unauthorizedStatus));
    }
}
