/*
Copyright 2022 Zheng Dayu
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at
    http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

package windowsupdate

import (
	"github.com/go-ole/go-ole"
	"github.com/pkg/errors"
	"time"
)

func okToErr(ok bool, t string) error {
	if !ok {
		return errors.Errorf("Not a %s", t)
	}
	return nil
}

func toIDispatchErr(result *ole.VARIANT, err error) (*ole.IDispatch, error) {
	if err != nil {
		return nil, err
	}
	return result.ToIDispatch(), nil
}

func toStringSliceErr(result *ole.VARIANT, err error) ([]string, error) {
	if err != nil {
		return nil, err
	}
	array := result.ToArray()
	if array == nil {
		return nil, nil
	}
	return array.ToStringArray(), nil
}

func toInt64Err(result *ole.VARIANT, err error) (int64, error) {
	if err != nil {
		return 0, err
	}
	return variantToInt64(result)
}

func toInt32Err(result *ole.VARIANT, err error) (int32, error) {
	if err != nil {
		return 0, err
	}
	return variantToInt32(result)
}

func toFloat64Err(result *ole.VARIANT, err error) (float64, error) {
	if err != nil {
		return 0, err
	}
	return variantToFloat64(result)
}

func toFloat32Err(result *ole.VARIANT, err error) (float32, error) {
	if err != nil {
		return 0, err
	}
	return variantToFloat32(result)
}

func toStringErr(result *ole.VARIANT, err error) (string, error) {
	if err != nil {
		return "", err
	}
	return variantToString(result)
}

func toBoolErr(result *ole.VARIANT, err error) (bool, error) {
	if err != nil {
		return false, err
	}
	return variantToBool(result)
}

func toTimeErr(result *ole.VARIANT, err error) (*time.Time, error) {
	if err != nil {
		return nil, err
	}
	return variantToTime(result)
}

func variantToInt64(v *ole.VARIANT) (int64, error) {
	valueRaw := v.Value()
	if valueRaw == nil {
		return 0, nil
	}

	value, ok := valueRaw.(int64)
	return value, okToErr(ok, "int64")
}

func variantToInt32(v *ole.VARIANT) (int32, error) {
	valueRaw := v.Value()
	if valueRaw == nil {
		return 0, nil
	}
	value, ok := valueRaw.(int32)
	return value, okToErr(ok, "int32")
}

func variantToFloat64(v *ole.VARIANT) (float64, error) {
	valueRaw := v.Value()
	if valueRaw == nil {
		return 0, nil
	}

	value, ok := valueRaw.(float64)
	return value, okToErr(ok, "float64")
}

func variantToFloat32(v *ole.VARIANT) (float32, error) {
	valueRaw := v.Value()
	if valueRaw == nil {
		return 0, nil
	}

	value, ok := valueRaw.(float32)
	return value, okToErr(ok, "float32")
}

func variantToString(v *ole.VARIANT) (string, error) {
	valueRaw := v.Value()
	if valueRaw == nil {
		return "", nil
	}

	value, ok := valueRaw.(string)
	return value, okToErr(ok, "string")
}

func variantToBool(v *ole.VARIANT) (bool, error) {
	valueRaw := v.Value()
	if valueRaw == nil {
		return false, nil
	}

	value, ok := valueRaw.(bool)
	return value, okToErr(ok, "bool")
}

func variantToTime(v *ole.VARIANT) (*time.Time, error) {
	valueRaw := v.Value()
	if valueRaw == nil {
		return nil, nil
	}

	value, ok := valueRaw.(time.Time)
	return &value, okToErr(ok, "time")
}
